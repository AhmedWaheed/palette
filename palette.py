'''
Created on Jul 24, 2016

@author: Ahmed Waheed <ahmedwaheed.mail@gmail.com>
'''

import sys
import time
import Queue
import Tkinter
import win32gui
import threading

class Window:
    '''
        The GUI of the program; 
        A Tkinter form consisting of: 
            - TextBox holding the hexadecimal RGB color code
            - Label displaying the selected color
            - Instruction label denoting whether the script is picking colors or paused
    '''
    def __init__(self, exit_handler):
        '''
            Initializing the form elements
        '''
        self.pick = True
        self.form = Tkinter.Tk() 
        self.name = "Main Window"            
        self.form.title("Color Picker")
        self.form.geometry("280x60+300+300")
        self.form.resizable(0, 0)       
        self.form.colorText = Tkinter.StringVar() 
        self.form.colorcode_textbox = Tkinter.Entry(self.form)
        self.form.colorcode_textbox.grid(row = 0, column=0)
        self.form.colorLabel = Tkinter.Label(self.form, bg = "#aaa", width = 20)
        self.form.colorLabel.grid(row = 0, column = 1)
        self.form.instructionLabel = Tkinter.Label(self.form, width = 20, text = "Press 'Esc' to save this color")
        self.form.instructionLabel.grid(row = 1, column = 0)
        # Events and protocols
        self.form.protocol("WM_DELETE_WINDOW", exit_handler)   # Closing parent program thread on closing the window
        self.form.bind('<Escape>', self.pause)  #pause when pressing 'Esc' button
        self.form.bind('<Return>', self.resume) #Resume when pressing 'Enter' button
        
    def readNewColor(self, queue):
        '''
            read the new cursor location and update the form to display its color 
        '''
        if not self.pick:
            return 
        while queue.qsize(  ):
            try:
                color = queue.get(0)
                self.form.colorText.set(color)
                self.form.colorcode_textbox.delete(0, 100)
                self.form.colorcode_textbox.insert(0, color)
                self.form.colorLabel.config(bg = color)
            except Queue.Empty:
                pass
            
    def run(self):
        print "Starting " + self.name + "..."
        self.form.mainloop()
        
    def pause(self, event):
        self.pick = False
        self.form.instructionLabel.configure(text = "Click 'Enter' to resume")
        
    def resume(self, event):
        self.pick = True
        self.form.instructionLabel.configure(text = "Press 'Esc' to save this color")

class Program:
    '''
        The main program:
            - It contains an instance of the Tkinter form "Window", 
            - It initiates a thread keeps track of the cursor position and record the color of this position in the queue "color_queue"
            - It updates the form with the new color
    '''
    def __init__(self):
        self._window = Window(self.stop)
        self.color_queue = Queue.Queue()

        self.running = True
        self.cursor_position_thread = threading.Thread(target=self.read_cursor_position)
        self.cursor_position_thread.start(  )

        self.updateForm(  )

    def read_cursor_position(self):
        '''
            reads the cursor position and retrieve its color
        '''
        while self.running:
            time.sleep(0.001)
            hdc = win32gui.GetDC(0)
            flags, hcursor, (x,y) = win32gui.GetCursorInfo()
            color = win32gui.GetPixel(hdc, x , y)
            color = self.rgb_int_to_hex(color)
            self.color_queue.put(color)

    
    def updateForm(self):
        '''
            call the method of the window that updates its color
        '''
        self._window.readNewColor(self.color_queue)
        if not self.running:
            sys.exit(1)
        self._window.form.after(1, self.updateForm)
            
    def rgb_int_to_hex(self, color):
        '''
            convert the color from Integer to RGB and then to hexadecimal notation
        '''
        rgb_int = int(color)
        red =  rgb_int & 255
        green = (rgb_int >> 8) & 255
        blue =   (rgb_int >> 16) & 255
        return '#%02x%02x%02x' % (red, green, blue)
    
    def stop(self):
        self.running = False
        print "Closing " + self._window.name + "..."

_program = Program()
_program._window.run()
