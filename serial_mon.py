#!/usr/bin/env python3
"""Sermon - a simple, single-file, cross-platform serial monitor."""

import codecs
import json
import os
import platform
import queue
import re
import sys
import threading
import time
import tkinter as tk
import tkinter.messagebox as tk_msg
import webbrowser
from pathlib import Path
from tkinter import filedialog, ttk
from typing import Dict, List, Optional

import serial
from serial.tools import list_ports

APP_NAME = 'sermon'
SYSTEM = platform.system()
APP_DIR = Path(__file__).resolve().parent
BUNDLED_PREFS = APP_DIR / 'preferences.json'
LEGACY_YAML_PREFS = APP_DIR / 'preferences.yaml'   # pre-JSON versions stored presets here

POLL_MS = 30             # how often the UI drains data coming from the reader thread
PORT_SCAN_MS = 1500      # port list auto-refresh while disconnected
READ_TIMEOUT_S = 0.1     # lets the reader thread notice stop requests
LINE_POLL_S = 0.05       # how often the modem input lines are sampled
MAX_LINES = 10000        # older output is trimmed to keep the text widget responsive
HEX_PER_ROW = 16
PORT_SEP = ' — '    # "COM7 — CP2102 USB to UART"

FLOW_CONTROLS = {'none': 'None', 'rtscts': 'RTS/CTS', 'xonxoff': 'XON/XOFF', 'dsrdtr': 'DSR/DTR'}
INPUT_LINES = ('CTS', 'DSR', 'RI', 'CD')
LINE_ON, LINE_OFF = '#2ecc71', '#d0d0d0'

DEFAULT_PROFILE = {'name': 'default', 'baud_rate': 115200, 'data_bits': 8,
                   'stop_bits': 1, 'parity': 'N', 'flow_control': 'none'}


# ---------------------------------------------------------------------------
# Preferences
# ---------------------------------------------------------------------------

def user_config_dir() -> Path:
    if SYSTEM == 'Windows':
        base = Path(os.environ.get('APPDATA') or Path.home())
    elif SYSTEM == 'Darwin':
        base = Path.home() / 'Library' / 'Application Support'
    else:
        base = Path(os.environ.get('XDG_CONFIG_HOME') or Path.home() / '.config')
    return base / APP_NAME


PREFS_FILE = user_config_dir() / 'preferences.json'


def validate_profile(raw) -> Optional[dict]:
    """Return a normalised copy of a connection profile, or None if it is invalid."""
    if not isinstance(raw, dict):
        return None
    try:
        prof = {
            'name': str(raw['name']),
            'baud_rate': int(raw['baud_rate']),
            'data_bits': int(raw['data_bits']),
            'stop_bits': float(raw['stop_bits']),
            'parity': str(raw['parity']),
            'flow_control': str(raw.get('flow_control', 'none')),
        }
    except (KeyError, TypeError, ValueError):
        return None
    if (prof['baud_rate'] <= 0
            or prof['data_bits'] not in serial.Serial.BYTESIZES
            or prof['stop_bits'] not in serial.Serial.STOPBITS
            or prof['parity'] not in serial.Serial.PARITIES
            or prof['flow_control'] not in FLOW_CONTROLS):
        return None
    if prof['stop_bits'].is_integer():
        prof['stop_bits'] = int(prof['stop_bits'])
    return prof


def _read_json(path: Path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def _read_legacy_yaml(path: Path):
    """Old versions kept presets in preferences.yaml. Import them once if PyYAML is around."""
    if not path.exists():
        raise OSError('no legacy preferences')
    try:
        import yaml
    except ImportError:
        print(f'Found {path} but PyYAML is not installed; old presets were not imported.',
              file=sys.stderr)
        raise OSError('PyYAML not available')
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def load_preferences() -> dict:
    """User preferences, else presets from an old YAML file, else the bundled defaults."""
    sources = ((PREFS_FILE, _read_json), (LEGACY_YAML_PREFS, _read_legacy_yaml),
               (BUNDLED_PREFS, _read_json))
    for path, reader in sources:
        try:
            data = reader(path)
        except Exception:   # missing file, bad JSON/YAML, ...
            continue
        if not isinstance(data, dict):
            continue
        profiles = {}
        for key, raw in (data.get('connection_profiles') or {}).items():
            prof = validate_profile(raw)
            if prof:
                profiles[str(key)] = prof
            else:
                print(f'Ignoring invalid profile {key!r} in {path}', file=sys.stderr)
        if profiles:
            current = (data.get('current_settings') or {}).get('connection_profile')
            if current not in profiles:
                current = next(iter(profiles))
            return {'connection_profiles': profiles,
                    'current_settings': {'connection_profile': current}}
    return {'connection_profiles': {'default': dict(DEFAULT_PROFILE)},
            'current_settings': {'connection_profile': 'default'}}


def save_preferences(prefs: dict) -> None:
    try:
        PREFS_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(PREFS_FILE, 'w', encoding='utf-8') as f:
            json.dump(prefs, f, indent=2)
            f.write('\n')
    except OSError as e:
        print(f'Could not save preferences to {PREFS_FILE}: {e}', file=sys.stderr)


# ---------------------------------------------------------------------------
# Port discovery
# ---------------------------------------------------------------------------

def port_sort_key(name: str):
    """Natural sort, so COM2 comes before COM10 and ttyUSB2 before ttyUSB10."""
    m = re.search(r'(\d+)$', name)
    return (name[:m.start()] if m else name, int(m.group(1)) if m else -1)


def windows_registry_ports() -> List[str]:
    """COM port names straight from the registry.

    pyserial's comports() walks the whole device tree to collect descriptions and
    can take seconds on Windows (Bluetooth COM ports are a common culprit), so the
    names come from here and the descriptions are filled in later in the background.
    """
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'HARDWARE\DEVICEMAP\SERIALCOMM') as key:
            count = winreg.QueryInfoKey(key)[1]
            names = [str(winreg.EnumValue(key, i)[1]) for i in range(count)]
    except OSError:
        return []
    return sorted(names, key=port_sort_key)


def detailed_ports() -> Dict[str, str]:
    """{device: description} for every port pyserial can see. May be slow on Windows."""
    ports = {}
    for p in list_ports.comports(include_links=True):
        desc = p.description if p.description not in (None, '', 'n/a', p.device) else ''
        ports[p.device] = desc
    return ports


# ---------------------------------------------------------------------------
# Serial I/O (runs on a worker thread; never touches Tk)
# ---------------------------------------------------------------------------

def read_input_lines(ser) -> Optional[Dict[str, bool]]:
    """CTS/DSR/RI/CD, or None when the port has no modem lines.

    Ptys, socat bridges, socket:// URLs and some Bluetooth ports fail here even though
    reading works, so a failure only means "unsupported"; unplugging is detected by read().
    """
    try:
        return {'CTS': ser.cts, 'DSR': ser.dsr, 'RI': ser.ri, 'CD': ser.cd}
    except Exception:
        return None


def reader_loop(ser, stop: threading.Event, q: queue.Queue) -> None:
    """Push ('data', bytes), ('lines', dict | None) and ('error', str) messages into q."""
    last_lines: Optional[Dict[str, bool]] = None
    lines_supported = True
    next_line_poll = 0.0
    while not stop.is_set():
        try:
            data = ser.read(ser.in_waiting or 1)
        except Exception as e:   # unplugged device, driver error, closed port...
            if not stop.is_set():
                q.put(('error', str(e) or e.__class__.__name__))
            return
        if data:
            q.put(('data', data))
        now = time.monotonic()
        if lines_supported and now >= next_line_poll:
            next_line_poll = now + LINE_POLL_S
            lines = read_input_lines(ser)
            lines_supported = lines is not None
            if lines != last_lines:
                last_lines = lines
                q.put(('lines', lines))


# ---------------------------------------------------------------------------
# Output formatting (pure, so it can be tested without a display)
# ---------------------------------------------------------------------------

class OutputFormatter:
    """Turns raw bytes into display text, keeping state across chunk boundaries."""

    def __init__(self):
        self.reset()

    def reset(self):
        self._decoder = codecs.getincrementaldecoder('utf-8')(errors='replace')
        self._pending_cr = False
        self._hex_col = 0

    def format(self, data: bytes, mode: str) -> str:
        if mode == 'hex':
            return self._hex(data)
        if mode == 'bytes':
            return ''.join(repr(bytes([b]))[2:-1] + ('\n' if b == 0x0A else '') for b in data)
        return self._text(data)

    def _text(self, data: bytes) -> str:
        text = self._decoder.decode(data)
        if self._pending_cr and text.startswith('\n'):
            text = text[1:]           # second half of a \r\n split across two reads
        if text:
            self._pending_cr = text.endswith('\r')
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        # Drop other control characters (except tab) that render as garbage in Tk.
        return ''.join(c for c in text if c >= ' ' or c in '\n\t')

    def _hex(self, data: bytes) -> str:
        parts = []
        for b in data:
            parts.append(f'{b:02x}')
            self._hex_col += 1
            if b == 0x0A or self._hex_col >= HEX_PER_ROW:
                parts.append('\n')
                self._hex_col = 0
            else:
                parts.append(' ')
        return ''.join(parts)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class SerialMon(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title('Serial Monitor')
        self.minsize(900, 600)

        self.ser = None
        self.read_thread = None
        self.read_stop = threading.Event()
        self.q = queue.Queue()
        self.formatter = OutputFormatter()

        self.port_names = []            # current port list
        self.port_desc = {}             # device -> description
        self.port_labels = {}           # combobox label -> device
        self.desc_fetched = set()       # ports we already asked descriptions for
        self.desc_thread = None

        self.preferences = load_preferences()
        self.available_profiles = self.preferences['connection_profiles']
        self.current_settings = tk.StringVar(
            value=self.preferences['current_settings']['connection_profile'])

        self._build_menu()
        self._build_widgets()

        self.handle_setting_change()
        self.set_connected_ui(False)
        self.protocol('WM_DELETE_WINDOW', self.on_close)

        self.after(POLL_MS, self.poll_queue)
        self.after(10, self.port_scan_tick)

    # -- layout ------------------------------------------------------------

    def _build_menu(self):
        self.menu_bar = tk.Menu(self)
        self.preferences_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label='Preferences', menu=self.preferences_menu)

        self.serial_port_settings_menu = tk.Menu(self.preferences_menu, tearoff=0)
        self.serial_port_settings_menu.add_command(label='Save preset', command=self.show_save_preset_pop)
        self.serial_port_settings_menu.add_separator()
        for name in self.available_profiles:
            self._add_profile_menu_entry(name)
        self.preferences_menu.add_cascade(label='Serial Port Settings', menu=self.serial_port_settings_menu)

        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.help_menu.add_command(label='About', command=self.show_about)
        self.menu_bar.add_cascade(label='Help', menu=self.help_menu)
        self.config(menu=self.menu_bar)

    def _add_profile_menu_entry(self, name):
        self.serial_port_settings_menu.add_radiobutton(
            label=name, value=name, variable=self.current_settings,
            command=self.handle_setting_change)

    def _build_widgets(self):
        pad = {'padx': 4, 'pady': 2}

        # Row 0: device, connection settings, control lines
        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky='we', **pad)

        self.devices_frame = ttk.Labelframe(top, text='Device')
        self.devices_frame.pack(side='left', fill='y', **pad)
        self.device_select = ttk.Combobox(self.devices_frame, width=34)
        self.device_select.grid(row=0, column=0, columnspan=2, sticky='we', **pad)
        self.device_connect = ttk.Button(self.devices_frame, text='Connect', command=self.handle_device_connection)
        self.device_connect.grid(row=1, column=0, sticky='we', **pad)
        self.devices_refresh = ttk.Button(self.devices_frame, text='Refresh', command=self.refresh_devices)
        self.devices_refresh.grid(row=1, column=1, sticky='we', **pad)

        self.settings_frame = ttk.Labelframe(top, text='Connection settings')
        self.settings_frame.pack(side='left', fill='y', **pad)
        bauds = [str(b) for b in serial.Serial.BAUDRATES[::-1]]
        self.baudrate_select = self._setting(0, 0, 'Baud', bauds, editable=True)  # custom rates allowed
        self.parity_select = self._setting(1, 0, 'Parity', list(serial.Serial.PARITIES))
        self.bytesize_select = self._setting(0, 2, 'Data bits', [str(b) for b in serial.Serial.BYTESIZES[::-1]])
        self.stopbits_select = self._setting(1, 2, 'Stop bits', [str(s) for s in serial.Serial.STOPBITS])
        self.flow_select = self._setting(2, 0, 'Flow', list(FLOW_CONTROLS.values()))
        self.flow_select.bind('<<ComboboxSelected>>', lambda e: self.update_control_line_state())

        self.lines_frame = ttk.Labelframe(top, text='Control lines')
        self.lines_frame.pack(side='left', fill='y', **pad)
        self.dtr_var = tk.BooleanVar(value=True)
        self.rts_var = tk.BooleanVar(value=True)
        self.dtr_check = ttk.Checkbutton(self.lines_frame, text='DTR', variable=self.dtr_var, command=self.on_dtr_toggle)
        self.rts_check = ttk.Checkbutton(self.lines_frame, text='RTS', variable=self.rts_var, command=self.on_rts_toggle)
        self.dtr_check.grid(row=0, column=0, sticky='w', **pad)
        self.rts_check.grid(row=0, column=1, sticky='w', **pad)
        self.break_btn = ttk.Button(self.lines_frame, text='Break', width=6, command=self.send_break)
        self.break_btn.grid(row=0, column=2, columnspan=2, sticky='we', **pad)
        self.line_leds = {}
        for i, name in enumerate(INPUT_LINES):
            led = tk.Label(self.lines_frame, text=name, width=4, relief='groove', bg=LINE_OFF, fg='#000000')
            led.grid(row=1, column=i, **pad)
            self.line_leds[name] = led

        # Row 1: send, output format, output actions
        mid = ttk.Frame(self)
        mid.grid(row=1, column=0, sticky='we', **pad)

        self.send_frame = ttk.Labelframe(mid, text='Send')
        self.send_frame.pack(side='left', fill='both', expand=True, **pad)
        self.send_entry = ttk.Entry(self.send_frame)
        self.send_entry.pack(side='left', fill='x', expand=True, **pad)
        self.send_entry.bind('<Return>', self.send_data)
        self.send_btn = ttk.Button(self.send_frame, text='Send', command=self.send_data)
        self.send_btn.pack(side='left', **pad)
        self.send_cr_value = tk.BooleanVar()
        self.send_lf_value = tk.BooleanVar()
        self.send_cr_check = ttk.Checkbutton(self.send_frame, text='CR', variable=self.send_cr_value)
        self.send_lf_check = ttk.Checkbutton(self.send_frame, text='LF', variable=self.send_lf_value)
        self.send_cr_check.pack(side='left', **pad)
        self.send_lf_check.pack(side='left', **pad)

        self.output_format_frame = ttk.Labelframe(mid, text='Output format')
        self.output_format_frame.pack(side='left', fill='y', **pad)
        self.output_format_var = tk.StringVar(value='txt')
        for value in ('txt', 'hex', 'bytes'):
            ttk.Radiobutton(self.output_format_frame, text=value, value=value, variable=self.output_format_var,
                            command=self.formatter.reset).pack(side='left', **pad)

        self.output_btn_frame = ttk.Labelframe(mid, text='Output text')
        self.output_btn_frame.pack(side='left', fill='y', **pad)
        ttk.Button(self.output_btn_frame, text='Clear', command=self.output_clear).pack(side='left', **pad)
        ttk.Button(self.output_btn_frame, text='Copy', command=self.output_copy_to_clipboard).pack(side='left', **pad)
        ttk.Button(self.output_btn_frame, text='Save to file', command=self.output_save_to_file).pack(side='left', **pad)

        # Row 2: output
        self.output_frame = ttk.Frame(self)
        self.output_frame.grid(row=2, column=0, sticky='nsew', **pad)
        self.output_scrollbar = ttk.Scrollbar(self.output_frame)
        self.output_text = tk.Text(self.output_frame, yscrollcommand=self.output_scrollbar.set,
                                   state='disabled', wrap='char', undo=False)
        self.output_scrollbar.config(command=self.output_text.yview)
        self.output_text.grid(row=0, column=0, sticky='nsew')
        self.output_scrollbar.grid(row=0, column=1, sticky='ns')
        self.output_frame.grid_columnconfigure(0, weight=1)
        self.output_frame.grid_rowconfigure(0, weight=1)
        self.output_text.tag_configure('tx', foreground='#1f6feb')
        self.output_text.tag_configure('info', foreground='#b03a2e')

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

    def _setting(self, row, col, label, values, editable=False):
        ttk.Label(self.settings_frame, text=label).grid(row=row, column=col, sticky='w', padx=4, pady=2)
        box = ttk.Combobox(self.settings_frame, values=values, width=9,
                           state='normal' if editable else 'readonly')
        box.grid(row=row, column=col + 1, sticky='we', padx=4, pady=2)
        return box

    # -- dialogs -------------------------------------------------------------

    def show_about(self):
        about_popup = tk.Toplevel(self)
        about_popup.title('About')
        tk.Label(about_popup, text='This is Sermon.\n\nIt is licensed under the GPL-v3 license.\n').pack(padx=12, pady=6)
        link = tk.Label(about_popup, text='https://github.com/picatostas/sermon', fg='blue', cursor='hand2')
        link.pack(padx=12, pady=6)
        link.bind('<Button-1>', lambda e: webbrowser.open_new(e.widget.cget('text')))

    def show_save_preset_pop(self):
        pop = tk.Toplevel(self)
        pop.title('Save preset')
        pop.transient(self)
        ttk.Label(pop, text='Enter the name:').grid(row=0, column=0, columnspan=2, padx=5, pady=5)
        entry = ttk.Entry(pop)
        entry.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        def ok(event=None):
            if self.save_preset(entry.get().strip()):
                pop.destroy()

        ttk.Button(pop, text='Save', command=ok).grid(row=2, column=0, padx=5, pady=5)
        ttk.Button(pop, text='Cancel', command=pop.destroy).grid(row=2, column=1, padx=5, pady=5)
        pop.bind('<Return>', ok)
        pop.bind('<Escape>', lambda e: pop.destroy())
        entry.focus_set()
        pop.grab_set()

    def save_preset(self, name) -> bool:
        if not name:
            tk_msg.showwarning(title='Save preset failed', message='Name cannot be empty')
            return False
        try:
            profile = self.read_settings()
        except ValueError as e:
            tk_msg.showwarning(title='Save preset failed', message=str(e))
            return False
        profile['name'] = name
        if name not in self.available_profiles:
            self._add_profile_menu_entry(name)
        self.available_profiles[name] = profile
        self.current_settings.set(name)
        return True

    # -- settings ------------------------------------------------------------

    def handle_setting_change(self):
        profile = self.available_profiles.get(self.current_settings.get())
        if profile is None:
            name, profile = next(iter(self.available_profiles.items()))
            self.current_settings.set(name)
        self.baudrate_select.set(str(profile['baud_rate']))
        self.parity_select.set(profile['parity'])
        self.bytesize_select.set(str(profile['data_bits']))
        self.stopbits_select.set(str(profile['stop_bits']))
        self.flow_select.set(FLOW_CONTROLS[profile.get('flow_control', 'none')])
        self.update_control_line_state()

    def flow_control(self) -> str:
        label = self.flow_select.get()
        return next((k for k, v in FLOW_CONTROLS.items() if v == label), 'none')

    def read_settings(self) -> dict:
        """Current widget values as a validated profile. Raises ValueError."""
        profile = validate_profile({
            'name': self.current_settings.get(),
            'baud_rate': self.baudrate_select.get().strip(),
            'data_bits': self.bytesize_select.get(),
            'stop_bits': self.stopbits_select.get(),
            'parity': self.parity_select.get(),
            'flow_control': self.flow_control(),
        })
        if profile is None:
            raise ValueError('Invalid connection settings.\nThe baud rate must be a positive whole number.')
        return profile

    # -- port list -------------------------------------------------------------

    def port_scan_tick(self):
        if self.ser is None:
            self.refresh_devices()
        self.after(PORT_SCAN_MS, self.port_scan_tick)

    def refresh_devices(self):
        if SYSTEM == 'Windows':
            names = windows_registry_ports()
            if names != self.port_names:
                self.port_names = names
                self.update_port_list()
            if not set(names) <= self.desc_fetched:
                self.fetch_port_details()
        else:
            self.fetch_port_details()   # comports() is quick on Linux and macOS

    def fetch_port_details(self):
        if self.desc_thread is not None and self.desc_thread.is_alive():
            return

        def worker():
            try:
                self.q.put(('ports', detailed_ports()))
            except Exception as e:
                print(f'Port enumeration failed: {e}', file=sys.stderr)

        self.desc_thread = threading.Thread(target=worker, daemon=True)
        self.desc_thread.start()

    def on_port_details(self, ports: Dict[str, str]):
        self.desc_fetched.update(ports)
        self.port_desc.update(ports)
        names = self.port_names if SYSTEM == 'Windows' else sorted(ports, key=port_sort_key)
        if SYSTEM == 'Windows':
            self.desc_fetched.update(names)   # don't keep asking for ports without a description
        self.port_names = names
        self.update_port_list()

    def update_port_list(self):
        labels = {}
        for name in self.port_names:
            desc = self.port_desc.get(name)
            labels[f'{name}{PORT_SEP}{desc}' if desc else name] = name
        if labels == self.port_labels:
            return
        selected = self.selected_port()
        self.port_labels = labels
        self.device_select['values'] = list(labels)
        # keep the user's selection, upgrading it to the labelled version if we now have one
        match = next((lbl for lbl, dev in labels.items() if dev == selected), None)
        if match:
            self.device_select.set(match)
        elif not selected and labels:
            self.device_select.set(next(iter(labels)))

    def selected_port(self) -> str:
        text = self.device_select.get().strip()
        return self.port_labels.get(text, text.split(PORT_SEP)[0].strip())

    # -- connection ------------------------------------------------------------

    def handle_device_connection(self):
        if self.ser is None:
            self.connect()
        else:
            self.disconnect()

    def connect(self):
        port = self.selected_port()
        if not port:
            tk_msg.showwarning(title='Devices', message='No device selected.\nPlease refresh, select or reconnect.')
            return
        try:
            settings = self.read_settings()
        except ValueError as e:
            tk_msg.showwarning(title='Connection settings', message=str(e))
            return

        flow = settings['flow_control']
        try:
            # serial_for_url also accepts URLs such as socket://host:port or rfc2217://
            ser = serial.serial_for_url(port, do_not_open=True)
            ser.baudrate = settings['baud_rate']
            ser.bytesize = settings['data_bits']
            ser.parity = settings['parity']
            ser.stopbits = settings['stop_bits']
            ser.timeout = READ_TIMEOUT_S
            ser.write_timeout = 2
            ser.rtscts = flow == 'rtscts'
            ser.xonxoff = flow == 'xonxoff'
            ser.dsrdtr = flow == 'dsrdtr'
            # Applied when the port opens. Untick DTR first to avoid resetting Arduino-style boards.
            ser.dtr = self.dtr_var.get()
            ser.rts = self.rts_var.get()
            ser.open()
        except (serial.SerialException, ValueError, OSError) as e:
            tk_msg.showerror(title='Devices', message=f"Couldn't connect to {port}:\n{e}")
            return

        self.ser = ser
        self.formatter.reset()
        self.read_stop = threading.Event()
        self.read_thread = threading.Thread(target=reader_loop, args=(ser, self.read_stop, self.q), daemon=True)
        self.read_thread.start()
        self.set_connected_ui(True)
        parity = settings['parity']
        self.title(f"Serial Monitor — {port} @ {settings['baud_rate']} "
                   f"{settings['data_bits']}{parity}{settings['stop_bits']}")

    def disconnect(self, error=None):
        ser, self.ser = self.ser, None
        if ser is None:
            return
        self.read_stop.set()
        if self.read_thread is not None:
            self.read_thread.join(timeout=1.0)
            self.read_thread = None
        try:
            ser.close()
        except Exception:
            pass
        self.set_connected_ui(False)
        self.title('Serial Monitor')
        if error:
            self.output_append(f'\n[Disconnected: {error}]\n', 'info')
        self.refresh_devices()

    def set_connected_ui(self, connected: bool):
        on, off = ('normal', 'disabled') if connected else ('disabled', 'normal')
        self.device_connect['text'] = 'Disconnect' if connected else 'Connect'
        self.devices_refresh['state'] = off
        self.device_select['state'] = off
        for box in (self.parity_select, self.bytesize_select, self.stopbits_select, self.flow_select):
            box['state'] = 'disabled' if connected else 'readonly'
        self.baudrate_select['state'] = off
        for widget in (self.send_entry, self.send_btn, self.send_cr_check, self.send_lf_check, self.break_btn):
            widget['state'] = on
        self.output_text.configure(bg='#ffffff' if connected else '#eeeeee',
                                   fg='#000000' if connected else '#777777')
        if not connected:
            self.update_input_lines(None)
        self.update_control_line_state()

    # -- control lines ---------------------------------------------------------

    def update_control_line_state(self):
        """RTS/DTR can't be toggled by hand while hardware flow control owns them."""
        flow = self.flow_control()
        self.rts_check['state'] = 'disabled' if flow == 'rtscts' else 'normal'
        self.dtr_check['state'] = 'disabled' if flow == 'dsrdtr' else 'normal'

    def on_dtr_toggle(self):
        self._set_line('dtr', self.dtr_var.get())

    def on_rts_toggle(self):
        self._set_line('rts', self.rts_var.get())

    def _set_line(self, name, value):
        if self.ser is None:
            return   # applied on connect
        try:
            setattr(self.ser, name, value)
        except (serial.SerialException, OSError, NotImplementedError) as e:
            self.output_append(f'\n[Could not set {name.upper()}: {e}]\n', 'info')

    def send_break(self):
        if self.ser is None:
            return
        try:
            self.ser.send_break(duration=0.25)
        except (serial.SerialException, OSError, NotImplementedError) as e:
            self.output_append(f'\n[Break failed: {e}]\n', 'info')

    def update_input_lines(self, lines: Optional[Dict[str, bool]]):
        for name, led in self.line_leds.items():
            state = None if lines is None else lines.get(name)
            led.configure(bg=LINE_ON if state else LINE_OFF, fg='#000000' if state is not None else '#888888')

    # -- data flow -------------------------------------------------------------

    def poll_queue(self):
        chunks = []
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == 'data':
                    chunks.append(payload)
                elif kind == 'lines':
                    if self.ser is not None:
                        self.update_input_lines(payload)
                elif kind == 'ports':
                    self.on_port_details(payload)
                elif kind == 'error':
                    if chunks:
                        self.output_append(self.formatter.format(b''.join(chunks), self.output_format_var.get()))
                        chunks = []
                    if self.ser is not None:
                        self.disconnect(payload)
        except queue.Empty:
            pass
        if chunks:
            self.output_append(self.formatter.format(b''.join(chunks), self.output_format_var.get()))
        self.after(POLL_MS, self.poll_queue)

    def send_data(self, event=None):
        if self.ser is None:
            return
        text = self.send_entry.get()
        payload = text.encode('utf-8')
        if self.send_cr_value.get():
            payload += b'\r'
        if self.send_lf_value.get():
            payload += b'\n'
        try:
            self.ser.write(payload)
        except (serial.SerialException, OSError) as e:
            self.disconnect(e)
            return
        lead = '' if self.output_text.get('end-2c', 'end-1c') in ('', '\n') else '\n'
        self.output_append(f'{lead}--> {text}\n', 'tx')
        self.formatter.reset()   # start the reply on a fresh line/row

    def output_append(self, text: str, tag: Optional[str] = None):
        if not text:
            return
        at_bottom = self.output_text.yview()[1] >= 0.999
        self.output_text.configure(state='normal')
        self.output_text.insert('end', text, (tag,) if tag else ())
        excess = int(self.output_text.index('end-1c').split('.')[0]) - MAX_LINES
        if excess > 0:
            self.output_text.delete('1.0', f'{excess + 1}.0')
        self.output_text.configure(state='disabled')
        if at_bottom:          # don't yank the view while the user scrolls back
            self.output_text.see('end')

    def output_clear(self):
        self.output_text.configure(state='normal')
        self.output_text.delete('1.0', 'end')
        self.output_text.configure(state='disabled')

    def output_copy_to_clipboard(self):
        self.clipboard_clear()
        self.clipboard_append(self.output_text.get('1.0', 'end-1c'))

    def output_save_to_file(self):
        filename = filedialog.asksaveasfilename(title='Save log as', defaultextension='.log',
                                                filetypes=[('Log files', '*.log'), ('All', '*')])
        if filename:
            with open(filename, 'w', encoding='utf-8') as out_f:
                out_f.write(self.output_text.get('1.0', 'end-1c'))

    def on_close(self):
        self.disconnect()
        self.preferences['current_settings']['connection_profile'] = self.current_settings.get()
        save_preferences(self.preferences)
        self.destroy()


def main():
    SerialMon().mainloop()


if __name__ == '__main__':
    main()
