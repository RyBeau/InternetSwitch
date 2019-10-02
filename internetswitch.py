"""
Author: Ryan Beaumont
Compatability: Windows 10
"""
import subprocess, ctypes, sys, pickle, os
from tkinter import *

"""File name of the persitant file used to store the interfaces"""
FILENAME = 'stored.interfaces'
"""Global list of the interface objects"""
INTERFACE_OBJECTS = None

class NetworkInterface:
    """
    This class defines a Network Interface object. A Network Interface has a state,
    name and command line.
    """
    def __init__(self, name):
        """
        Initialising all of variables for the Network Interface. Then setting the 
        initial off state by call self._off()
        """
        self.state = True
        self.name = name
        self.command = 'netsh interface set interface ' + '"{}"'.format(name)        
        self._off()

    
    def _off(self):
        """This method uses the command line for the interface to disable the
        the Network Interface, It then updates the state of the adapter to False
        to reflect this change"""
        string = self.command + " Disable"
        subprocess.call(string)
        self.state = False
    
    def _on(self):
        """This method uses the command line for the interface to enable the
        the Network Interface, It then updates the state of the adapter to True
        to reflect this change"""
        string = self.command + " Enable"
        subprocess.call(string)
        self.state = True
        
class Setup:
    """
    This class defines the Setup window which the user will use in the case
    of first setup. From this window the select with Network Interfaces they would
    like to be able to control using the program.
    """
    def __init__(self, window):
        """
        Initialising the window and getting the list of network adapters installed
        on the system.
        """
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
        """
        This method uses the cmd command defined in cmd to receive a list of 
        Network Adapters from the cmd window. It then uses the data from the window
        and then extracts the names of the Network Adapters from the output of the
        cmd window. The names are returned in a list.
        """
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
        """
        This method creates the list of Network Adapter names to be turned into
        Network Adapter Objects as selected by the user. This then sends this list
        to self.store() to create the Network Adapter Objects.
        """
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
        """
        This method creates the Network Adapter Objects and the stores them
        in the global interface list for usage in the rest of the program.
        """
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
    """
    Begins the first setup process.
    """
    window = Tk()
    start = Setup(window)
    window.configure(background='Alice Blue',borderwidth=2,relief=SOLID)
    window.mainloop()

def re_run_setup(root):
    """
    This function reruns setup as by removing FILENAME and then restarting the 
    program.
    """
    try:
        os.remove(FILENAME)
    except:
        root.destroy()
        run_as_admin()
    else:
        root.destroy()
        run_as_admin()        
        
def check_admin():
    """
    Checks that the program is being run as admin or calls run_as_admin.
    """
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def load_interfaces():
    """
    Loads the interfaces stored in FILENAME
    """
    with open(FILENAME, 'rb') as INTERFACE_OBJECTS_file:
        interfaces = pickle.load(INTERFACE_OBJECTS_file)    
    return interfaces

def run_as_admin():
    """Runs the program as admin"""
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