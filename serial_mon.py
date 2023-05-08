import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as tk_msg
from tkinter import filedialog

import enum

import threading
import serial

import platform

import yaml
from jsonschema import validate

import webbrowser

class DevState(enum.Enum):
    NC = 0
    CONNECTED = 1


class SerialMon(tk.Tk):

    def __init__(self):
        super(SerialMon, self).__init__()
        self.title('Serial Monitor')
        self.conn_status = DevState.NC
        self.minsize(820, 600)
        self.resizable(True, True)

        # Load preferences from YAML file
        with open('preferences.yaml', 'r') as f:
            self.preferences = yaml.safe_load(f)

        # Load schema from YAML file
        with open('preferences-schema.yaml', 'r') as f:
            schema = yaml.safe_load(f)

        # Validate each connection profile against the schema
        for profile in self.preferences['connection_profiles'].values():
            validate(profile, schema['definitions']['connection_profile'])

        self.available_profiles = self.preferences['connection_profiles']
        self.current_settings = tk.StringVar(value=self.preferences['current_settings']['connection_profile'])
        # Create menu bar
        self.menu_bar = tk.Menu(self)

        # Create preferences menu
        self.preferences_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Preferences", menu=self.preferences_menu)

        # Create serial port settings menu
        self.serial_port_settings_menu = tk.Menu(self.preferences_menu, tearoff=0)
        self.serial_port_settings_menu.add_command(label="Save preset", command=self.show_save_preset_pop)
        for profile in self.available_profiles:
            self.serial_port_settings_menu.add_radiobutton(label=profile, value=profile, variable=self.current_settings,
                                                    command=self.handle_setting_change,
                                                    indicatoron=1, activebackground='gray')

        self.preferences_menu.add_cascade(label="Serial Port Settings", menu=self.serial_port_settings_menu)

        # Create About menu
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.help_menu.add_command(label="About", command=self.show_about)
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)
        self.config(menu=self.menu_bar)

        self.devices_frame = ttk.Labelframe(self, text="Devices")

        self.devices = self.get_devices()
        self.device_select = ttk.Combobox(self.devices_frame, values=self.devices)
        self.device_select.grid(row=0, column=0, sticky="WE", columnspan=2)
        self.device_connect = ttk.Button(self.devices_frame, text="Connect", command=self.handle_device_connection)
        self.device_connect.grid(row=1, column=0, sticky="E")
        self.devices_refresh = ttk.Button(self.devices_frame, text="Refresh", command=self.refresh_devices)
        self.devices_refresh.grid(row=1, column=1, sticky="E")

        self.devices_frame.grid(row=0, column=0, sticky="W")

        self.settings_frame = ttk.Labelframe(self, text="Connection settings")

        self.baudrate_select = ttk.Combobox(self.settings_frame, values=list(serial.SerialBase.BAUDRATES)[::-1], state="readonly")
        self.baudrate_label = tk.Label(self.settings_frame, text='Baudrate')
        self.baudrate_label.grid(row=0, column=0, sticky="WE")
        self.baudrate_select.grid(row=0, column=1, sticky="WE")

        self.parity_select = ttk.Combobox(self.settings_frame, values=list(serial.SerialBase.PARITIES), state="readonly")
        self.parity_label = tk.Label(self.settings_frame, text='Parity')
        self.parity_label.grid(row=1, column=0, sticky="WE")
        self.parity_select.grid(row=1, column=1, sticky="WE")

        self.bytesize_select = ttk.Combobox(self.settings_frame, values=list(serial.SerialBase.BYTESIZES)[::-1], state="readonly")
        self.bytesize_label = tk.Label(self.settings_frame, text='Bytesize')
        self.bytesize_label.grid(row=0, column=2, sticky="WE")
        self.bytesize_select.grid(row=0, column=3, sticky="WE")

        self.stopbits_select = ttk.Combobox(self.settings_frame, values=list(serial.SerialBase.STOPBITS), state="readonly")
        self.stopbits_label = tk.Label(self.settings_frame, text='Stopbits')
        self.stopbits_label.grid(row=1, column=2, sticky="WE")
        self.stopbits_select.grid(row=1, column=3, sticky="WE")

        self.settings_frame.grid(row=0, column=1, columnspan=2, sticky='NSw')

        self.send_frame = tk.LabelFrame(self, text='Send')
        self.send_entry = ttk.Entry(self.send_frame)
        self.send_entry.grid(row=0, column=0, sticky="w")
        self.send_btn = ttk.Button(self.send_frame, text="Send", command=self.send)
        self.send_btn.grid(row=0, column=1, sticky="w")
        self.send_cr_value = tk.BooleanVar()
        self.send_cr_check = tk.Checkbutton(self.send_frame, text='CR', variable=self.send_cr_value)
        self.send_lf_value = tk.BooleanVar()
        self.send_lf_check = tk.Checkbutton(self.send_frame, text='LF', variable=self.send_lf_value)
        self.send_cr_check.grid(row=0, column=2)
        self.send_lf_check.grid(row=0, column=3)

        self.send_entry.bind("<Return>", self.send)

        self.send_frame.grid(row=1, column=0, columnspan=2, sticky="w")

        self.output_frame = tk.Frame(self)

        self.output_scrollbar = tk.Scrollbar(self.output_frame)
        self.output_text = tk.Text(self.output_frame, yscrollcommand=self.output_scrollbar.set)

        self.read_buffer = ''
        # Add an interval variable to control the update frequency
        self.update_interval = 50  # in ms

        self.output_format_frame = ttk.Labelframe(self, text="Output format")
        self.output_format_var = tk.StringVar(self.output_format_frame, 'txt')

        self.output_format_txt = ttk.Radiobutton(self.output_format_frame, text='txt', variable=self.output_format_var, value='txt')
        self.output_format_hex = ttk.Radiobutton(self.output_format_frame, text='hex', variable=self.output_format_var, value='hex')
        self.output_format_bytes = ttk.Radiobutton(self.output_format_frame, text='bytes', variable=self.output_format_var, value='bytes')

        self.output_btn_frame = ttk.Labelframe(self, text="Output text")
        self.output_clear_btn = ttk.Button(self.output_btn_frame, text='Clear', command=self.output_clear)
        self.output_copy_btn = ttk.Button(self.output_btn_frame, text='Copy to clipboard', command=self.output_copy_to_clipboard)
        self.output_save_btn = ttk.Button(self.output_btn_frame, text='Save to file', command=self.output_save_to_file)
        self.output_text.configure(state="disabled", bg="#eeeeee", fg="#999999")

        for item in self.settings_frame.winfo_children():
            item['state'] = tk.NORMAL

        for item in self.send_frame.winfo_children():
            item['state'] = tk.DISABLED

        # Stop event/signal for the serial read thread
        self.read_stop_event = threading.Event()

        self.output_scrollbar.config(command=self.output_text.yview)
        self.output_text.grid(row=0, column=0, sticky="nswe")
        self.output_scrollbar.grid(row=0, column=1, sticky="nse")

        self.output_text.grid_rowconfigure(0, weight=1)
        self.output_scrollbar.grid_rowconfigure(0, weight=1)

        self.output_format_txt.grid(row=0, column=0, sticky="e")
        self.output_format_hex.grid(row=0, column=1, sticky="e")
        self.output_format_bytes.grid(row=0, column=2, sticky="e")

        self.output_clear_btn.grid(row=1, column=0, sticky="e")
        self.output_copy_btn.grid(row=1, column=1, sticky="e")
        self.output_save_btn.grid(row=1, column=2, sticky="e")

        self.output_format_frame.grid(row=1, column=1, sticky="nsew")

        self.output_btn_frame.grid(row=1, column=2, sticky="nse")

        self.output_frame.grid(row=2, column=0, columnspan=self.grid_size()[0], sticky="nsew")
        self.output_frame.grid_columnconfigure(0, weight=1)
        self.output_frame.grid_rowconfigure(0, weight=1)

        # Make the frames follow the window size
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        # apply the read profile
        self.handle_setting_change()
        # split the exit sequence so we can do some stuff before the application closes
        self.protocol("WM_DELETE_WINDOW", self.save_preferences)

    def show_about(self):
        # Show the about dialog
        about_popup = tk.Toplevel(self)
        about_text = f"This is Sermon. \n\nIt is licensed under the GPL-v3 license.\n\n"
        about_popup.title("About")
        about_label = tk.Label(about_popup, text=about_text)
        about_label.pack()
        link_label = tk.Label(about_popup, text="https://github.com/picatostas", fg="blue", cursor="hand2")
        link_label.pack()

        def callback(event):
            webbrowser.open_new(event.widget.cget("text"))

        link_label.bind("<Button-1>", callback)

    def handle_setting_change(self):
        profile_name = self.current_settings.get()
        profile = self.available_profiles[profile_name]
        print(f'Setting changed to: {profile}')
        self.parity_select.set(profile['parity'])
        self.baudrate_select.set(profile['baud_rate'])
        self.bytesize_select.set(profile['data_bits'])
        self.stopbits_select.set(profile['stop_bits'])

    def show_save_preset_pop(self):

        self.save_preset_pop = tk.Toplevel(self)
        self.save_preset_pop.title("Save preset")

        self.save_preset_label = tk.Label(self.save_preset_pop, text="Enter the name:")
        self.save_preset_label.grid(row=0, column=0, padx=5, pady=5)
        self.save_preset_entry = tk.Entry(self.save_preset_pop)
        self.save_preset_entry.grid(row=1, column=0, padx=5, pady=5)

        self.save_preset_ok_cancel_frame = tk.Frame(self.save_preset_pop)
        self.ok_button = tk.Button(self.save_preset_ok_cancel_frame, text="Save", command=self.save_preset_pop_ok)
        self.ok_button.grid(row=2, column=0, padx=5, pady=5)
        self.cancel_button = tk.Button(self.save_preset_ok_cancel_frame, text="Cancel", command=self.save_preset_pop_cancel)
        self.cancel_button.grid(row=2, column=1, padx=5, pady=5)
        self.save_preset_ok_cancel_frame.grid(row=2, column=0, columnspan=2)
        self.save_preset_entry.focus_set()

        self.save_preset_pop.bind("<Return>", self.save_preset_pop_ok)
        self.save_preset_pop.bind("<Escape>", self.save_preset_pop_cancel)


    def save_preset_pop_ok(self, event):

        user_input = self.save_preset_entry.get()

        new_profile = dict()
        new_profile['name'] = user_input
        new_profile['parity'] = self.parity_select.get()
        new_profile['baud_rate'] = int(self.baudrate_select.get())
        new_profile['data_bits'] = int(self.bytesize_select.get())
        new_profile['stop_bits'] = float(self.stopbits_select.get())

        has_empty_values = False
        for value in new_profile.values():
            if value == "":
                has_empty_values = True
                break

        print(new_profile)

        if user_input == "":
            tk_msg.showwarning(title="Save preset failed", message="Name cannot be empty")
        elif has_empty_values:
            tk_msg.showwarning(title="Save preset failed", message="One of more values are empty")
            self.save_preset_pop.destroy()
        else:
            # Save the new connection profiles to the preferences
            self.preferences['connection_profiles'][user_input] = new_profile
            self.serial_port_settings_menu.add_radiobutton(label=new_profile['name'], value=new_profile, variable=self.current_settings,
                                        command=self.handle_setting_change,
                                        indicatoron=1, activebackground='gray')
            self.save_preset_pop.destroy()

    def save_preset_pop_cancel(self,event):
        # Close the dialog without doing anything
        self.save_preset_pop.destroy()

    def save_preferences(self):
        # Serialize the updated dictionary to a YAML file
        with open('preferences.yaml', 'w') as f:
            yaml.dump(self.preferences, f)
        self.destroy()


    def output_clear(self):
        self.output_text.configure(state="normal")
        self.output_text.delete(1.0, tk.END)
        self.output_text.configure(state="disabled")

    def output_copy_to_clipboard(self):

        # Clear the clipboard
        self.clipboard_clear()

        # Append the contents to the clipboard
        self.clipboard_append(self.output_text.get(1.0, tk.END))

    def output_save_to_file(self):
        filename = filedialog.asksaveasfilename(title="Select a log:", initialdir="./", filetypes=[("Log files", "*.log"), ("All", "*")])
        if type(filename) is str:
            with open(filename, 'w') as out_f:
                out_f.write(self.output_text.get(1.0, tk.END))

    def get_devices(self):
        serial_devs = []
        system_platform = platform.system()

        # For some reason that I ignore, serial.tools.list_ports takes forever to run
        # in Win 11, so fetch the com ports from the winreg entry instead.
        if system_platform == 'Windows':
            import winreg
            i = 0
            # This can drop an exception maybe
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'HARDWARE\\DEVICEMAP\\SERIALCOMM')
            except:
                print("No serial ports found")
                return serial_devs

            i = 0
            # iter over the possible values
            while True:
                try:
                    serial_devs.append(winreg.EnumValue(key, i)[1])
                    i += 1
                except:
                    # end of entries
                    break

        # For Linux and Mac serial.tools works just fine
        if system_platform == 'Linux' or system_platform == 'Darwin':
            from serial.tools import list_ports
            ports = list(list_ports.comports(True))

            for port in ports:
                serial_devs.append(port.device)

        return serial_devs

    def refresh_devices(self):

        devices = self.get_devices()
        self.device_select['values'] = devices

    def handle_device_connection(self):
        print('Handle Device connection')
        if self.conn_status == DevState.NC:
            has_empty_values = False

            active_profile = dict()

            active_profile['parity'] = self.parity_select.get()
            active_profile['baud_rate'] = self.baudrate_select.get()
            active_profile['data_bits'] = self.bytesize_select.get()
            active_profile['stop_bits'] = self.stopbits_select.get()

            for value in active_profile.values():
                if value == "":
                    has_empty_values = True
                    break

            if has_empty_values:
                self.current_settings.set('default')
            self.handle_setting_change()


            _baudrate = int(self.baudrate_select.get())

            _parity = self.parity_select.get()

            _bytesize = int(self.bytesize_select.get())

            _stopbits = float(self.stopbits_select.get())

            _port = self.device_select.get()

            if _port == '':
                tk_msg.showwarning(title='Devices', message='No device found nor selected, \nplease refresh, select or reconnect')
            else:
                try:
                    self.ser = serial.Serial(port=_port,
                                             baudrate=_baudrate,
                                             bytesize=_bytesize,
                                             parity=_parity,
                                             stopbits=_stopbits)
                    print(f'Serial config:')
                    print(f'\tport: {_port} baudrate: {_baudrate} parity: {serial.PARITY_NAMES[_parity]}')
                    print(f'\tbytesize: {_bytesize} stopbits: {_stopbits}')

                    self.device_connect['text'] = 'Disconnect'
                    self.conn_status = DevState.CONNECTED
                    self.output_text.configure(fg='#000000', bg='#ffffff')
                    self.read_stop_event.clear()

                    # Start the thread each time as apparently a thread cannot be started after being
                    # terminated
                    read_thread = threading.Thread(target=self.read_thread_target, args=[self.read_stop_event])
                    # set the process as daemon so it dies if the main process terminates
                    read_thread.daemon = True
                    read_thread.start()

                    self.devices_refresh['state'] = tk.DISABLED
                    self.device_select['state'] = tk.DISABLED

                    for item in self.settings_frame.winfo_children():
                        item['state'] = tk.DISABLED

                    for item in self.send_frame.winfo_children():
                        item['state'] = tk.NORMAL

                except:
                    tk_msg.showerror(title='Devices', message=f'Couldn\'t connect to device {_port}, \nplease refresh or reconnect')

        elif self.conn_status == DevState.CONNECTED:

            if self.ser is not None:
                if self.ser.is_open:
                    print(f'Disconnected from device: {self.ser.port}')
                    self.device_connect['text'] = 'Connect'
                    self.conn_status = DevState.NC
                    self.output_text.configure(bg="#eeeeee", fg="#999999")
                    self.ser.close()
                    self.read_stop_event.set()
                    self.devices_refresh['state'] = tk.NORMAL
                    self.device_select['state'] = tk.NORMAL

                    for item in self.settings_frame.winfo_children():
                        item['state'] = tk.NORMAL

                    for item in self.send_frame.winfo_children():
                        item['state'] = tk.DISABLED

    def read_thread_target(self, event):
        print('Start reading thread')
        while True:
            # stop event, if received, die.
            if event.is_set():
                print('Received termination event')
                exit(0)

            if self.conn_status == DevState.CONNECTED:
                data_str = self.ser.read().decode('utf-8')
                self.read_buffer += data_str
                self.after(self.update_interval, self.update_text_box)

    def update_text_box(self):
        self.output_append(self.read_buffer, prefix='')
        self.read_buffer = ''


    def send(self, event=None):
        send_str = str(self.send_entry.get())
        print(f"send: {send_str}")
        self.output_append(send_str, prefix='\n--> ', new_line=True)
        self.output_append('\n', prefix='<-- ', new_line=False)
        send_bytes = bytearray()
        send_bytes.extend(send_str.encode())
        self.ser.write(send_bytes)
        if self.send_cr_value.get():
            self.ser.write(serial.CR)
        if self.send_lf_value.get():
            self.ser.write(serial.LF)

    def output_append(self, data_in: str, prefix: str = '', new_line : bool = False):

        data_str = ''

        if self.output_format_var.get() == 'txt':
            data_str += data_in
        elif self.output_format_var.get() == 'hex':
            for i in data_in:
                data_str += f'0x{ord(i):02x} '
        elif self.output_format_var.get() == 'bytes':
            for i in data_in:
                data_str += str(bytes(i, 'utf-8')) + ' '

        if new_line:
            data_str += '\n'

        out_str = prefix + data_str
        self.output_text.configure(state="normal")
        self.output_text.insert("end", out_str)
        self.output_text.see(tk.END)
        self.output_text.configure(state="disabled")


def main():
    serial_mon = SerialMon()
    serial_mon.mainloop()


if __name__ == '__main__':
    main()
