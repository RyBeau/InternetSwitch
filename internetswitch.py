import subprocess, ctypes, sys, pickle, os
from tkinter import *

FILENAME = 'stored.interfaces'
INTERFACE_OBJECTS = None

class NetworkInterface:
    def __init__(self, name):
        self.state = True
        self.name = name
        self.command = 'netsh interface set interface ' + '"{}"'.format(name)        
        self._off()

    
    def _off(self):
        string = self.command + " Disable"
        subprocess.call(string)
        self.state = False
    
    def _on(self):
        string = self.command + " Enable"
        subprocess.call(string)
        self.state = True
        
class Setup:
    def __init__(self, window):
        self.window = window
        self.network_interfaces = self.get_interfaces()
        self.title = Label(window, text="Select the Network Interfaces to Control",
              bg="Alice Blue")
        self.title.pack(anchor=N,padx = 5)
        chk_frame = Frame(window,bg="Alice Blue", borderwidth=1,relief=GROOVE)
        self.checkboxes = []
        for item in self.network_interfaces:
            var = IntVar()
            chk = Checkbutton(chk_frame, text=item, variable=var,bg="Alice Blue",
                         activebackground="Alice Blue")
            chk.pack(anchor = W,padx = 5)
            self.checkboxes.append(var)
        chk_frame.pack(anchor = N,padx = 5)
        confirm = Button(window, text="Confirm", command=self.set_interfaces,
                         width=10,bg="Alice Blue",
                         activebackground="Alice Blue").pack(pady=5)
        
    def get_interfaces(self):
        cmd = 'netsh interface show interface'
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True)
        (output, _) = p.communicate()
        output_str = str(output)
        output_list = []
        for item in output_str.split("\\r\\n"):
            output_list.append(item)
        network_interfaces = []
        for item in output_list:
            if "Enabled" in item or "Disabled" in item:
                line = item.split("        ")
                network_interfaces.append(line[-1])
        network_interfaces.sort()
        return network_interfaces
    

    def set_interfaces(self):
        values = []
        for chk in self.checkboxes:
            values.append(chk.get())
        chosen_interfaces = []
        for i in range(len(values)):
            if values[i] == 1:
                chosen_interfaces.append(self.network_interfaces[i])
        if len(chosen_interfaces) > 0:
            self.store(chosen_interfaces)
        else:
            self.title.configure(text="Select an interface needs to be selected")
       
    
    def store(self,chosen_interfaces):
        interface_objects = []
        for item in chosen_interfaces:
            interface_objects.append(NetworkInterface(item))
        global INTERFACE_OBJECTS
        INTERFACE_OBJECTS = interface_objects
        self.window.destroy()
        
        
class GUI:
    def __init__(self, root, current):
        self.root = root
        self.interface_objects = INTERFACE_OBJECTS
        self.current_interface = current
        button_frame_1 = Frame(root,bg="Alice Blue")
        self.status = Label(button_frame_1,text=self.current_interface.name,
                            bg="Alice Blue")
        self.status.pack()
        button_list = []
        for item in self.interface_objects:
            button_list.append(Button(button_frame_1,text=item.name,
                                    command= lambda x = item: self.switch(x)
                                    ,width=10,bg="Alice Blue",
                                    activebackground="Alice Blue").pack(pady=5))
        setup_button = Button(button_frame_1,text="Setup",
                                    command= lambda: re_run_setup(self.root)
                                    ,width=10,bg="Alice Blue",
                                    activebackground="Alice Blue").pack(pady=5)
        exit_button = Button(button_frame_1,text="Exit",
                             command= self.quit,bg="Alice Blue",
                             activebackground="Alice Blue",
                             width=10).pack(pady=5)
        
        button_frame_1.pack(padx=5)

    def switch(self, interface):
        if interface != self.current_interface:
            self.current_interface._off()
            interface._on()
            self.current_interface = interface
            self.status.configure(text= self.current_interface.name)

    def store_interfaces(self):
        with open(FILENAME, 'wb+') as INTERFACE_OBJECTS_file:
            pickle.dump(self.interface_objects, INTERFACE_OBJECTS_file)

    def quit(self):
        self.store_interfaces()
        for item in self.interface_objects :
            if item.state is False:
                item._on()
        self.root.destroy()
  
    
def main():
    on = INTERFACE_OBJECTS[0]
    on._on()
    root = Tk()
    start = GUI(root, on)
    root.geometry("-0-30")
    root.overrideredirect(1)
    root.configure(background='Alice Blue',borderwidth=2,relief=SOLID)
    root.mainloop()

def first_setup():
    window = Tk()
    start = Setup(window)
    window.configure(background='Alice Blue',borderwidth=2,relief=SOLID)
    window.mainloop()

def re_run_setup(root):
    try:
        os.remove(FILENAME)
    except:
        root.destroy()
        run_as_admin()
    else:
        root.destroy()
        run_as_admin()        
        
def check_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def load_interfaces():
    with open(FILENAME, 'rb') as INTERFACE_OBJECTS_file:
        interfaces = pickle.load(INTERFACE_OBJECTS_file)    
    return interfaces

def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)

if __name__ == "__main__":
    if check_admin():
        if os.path.exists(FILENAME):
            INTERFACE_OBJECTS = load_interfaces()
            main()            
        else:
            first_setup()
            main()            
    else:
        run_as_admin()