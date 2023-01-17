import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as tk_msg
from tkinter import filedialog

import enum

import threading
import serial

import platform


class DevState(enum.Enum):
    NC = 0
    CONNECTED = 1


class SerialMon(tk.Tk):

    def __init__(self):
        super(SerialMon, self).__init__()
        self.title('Serial Monitor')
        self.conn_status = DevState.NC
        self.resizable(False, False)
        self.devices_frame = ttk.Labelframe(self, text="Devices")

        self.devices = self.get_devices()
        self.device_select = ttk.Combobox(self.devices_frame, values=self.devices, state="readonly")
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

        self.settings_frame.grid(row=0, column=2, sticky='NSw')

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
        self.output_text.grid(row=0, column=0, sticky="we")
        self.output_scrollbar.grid(row=0, column=1, sticky="nse")

        self.output_format_txt.grid(row=0, column=0, sticky="e")
        self.output_format_hex.grid(row=0, column=1, sticky="e")
        self.output_format_bytes.grid(row=0, column=2, sticky="e")

        self.output_clear_btn.grid(row=1, column=0, sticky="e")
        self.output_copy_btn.grid(row=1, column=1, sticky="e")
        self.output_save_btn.grid(row=1, column=2, sticky="e")

        self.output_format_frame.grid(row=0, column=1, sticky="nsew")

        self.output_btn_frame.grid(row=1, column=2, sticky="nsew")

        self.output_frame.grid(row=2, column=0, columnspan=self.grid_size()[0], sticky="nsew")
        self.output_frame.grid_columnconfigure(0, weight=1)

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

            _baudrate = 921600 if self.baudrate_select.get() == '' else int(self.baudrate_select.get())

            _parity = serial.PARITY_NONE if self.parity_select.get() == '' else self.parity_select.get()

            _bytesize = serial.EIGHTBITS if self.bytesize_select.get() == '' else int(self.bytesize_select.get())

            _stopbits = serial.STOPBITS_ONE if self.stopbits_select.get() == '' else float(self.stopbits_select.get())

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
                self.output_append(data_str, prefix='')

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
