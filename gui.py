import wx
import test
import os
from cert_type_enum import cert_type

# Define the size of the application window
APP_SIZE_X = 500
APP_SIZE_Y = 400

class gui(wx.Dialog):

    # Initialize instance variables
    param1 = False
    param2 = False
    error_1_text = None
    error_1_bool = False
    save_directory = None
    roster_and_grades = None
    doc_template = cert_type.DEFAULT.value
     
    def __init__(self, parent, id, title):
        # Call the superclass constructor to create the dialog window
        wx.Dialog.__init__(self, parent, id, title,
                           size=(APP_SIZE_X, APP_SIZE_Y))                   

        # Create two vertical box sizers for the two labels and their input controls
        vbox1 = wx.BoxSizer(wx.VERTICAL)
        label1 = wx.StaticText(self, label='Select Roster and Grades:')
        label1.SetPosition((30,20))
        vbox1.Add(label1, 0, wx.ALL, 5)

        vbox2 = wx.BoxSizer(wx.VERTICAL)
        label2 = wx.StaticText(self, label='Output Directory:')
        label2.SetPosition((30,70))
        vbox2.Add(label2, 0, wx.ALL, 5)

        # Create a horizontal box sizer to contain the two vertical box sizers
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(vbox1, 0, wx.EXPAND | wx.ALL, 5)
        hbox.Add(vbox2, 0, wx.EXPAND | wx.ALL, 5)

        # Add buttons and radio box to the dialog
        radio_list = ['Default', 'SBA', 'DOIU','NOAA']
        self.file_picker = wx.FilePickerCtrl(self, 1, message='Select Roster and Grades', pos=(30,40), style=wx.FLP_USE_TEXTCTRL)
        self.directory_picker = wx.DirPickerCtrl(self, 2, message='Output Directory', pos=(30,90), style=wx.DIRP_USE_TEXTCTRL)
        wx.Button(self, 3, 'Create Certificates', (350, 300))
        self.rbox = wx.RadioBox(self, 4, 'Certificate Type', choices=radio_list, pos=(30,130))

        # Set the initial sizes of the input controls
        self.file_picker.SetInitialSize((400, -1))
        self.directory_picker.SetInitialSize((400, -1))

        # Add the input controls to the vertical box sizers
        vbox1.Add(self.file_picker, 0, wx.ALL, 5)
        vbox2.Add(self.directory_picker, 0, wx.ALL, 5)

        # Bind events to the buttons and radio box
        self.Bind(wx.EVT_FILEPICKER_CHANGED, self.file_select, id=1)
        self.Bind(wx.EVT_DIRPICKER_CHANGED, self.output_select, id=2)
        self.Bind(wx.EVT_BUTTON, self.run, id=3)
        self.rbox.Bind(wx.EVT_RADIOBOX, self.doc_select, id=4)

        # Center and show the dialog
        self.Centre()
        self.ShowModal()
        self.Destroy()

    def shorten(file_path):
        # Split file_path into a list and return only the last two elements joined by a '/'
        return '/'.join(file_path.split('/')[-2:])

    def file_select(self, event):
        # When the user selects a file with the file picker, update the instance variables and display a relative file
        abs_path = self.file_picker.GetPath()
        gui.roster_and_grades = abs_path
        cwd = os.getcwd()
        rel_path = os.path.relpath(abs_path, cwd)
        self.file_picker.SetPath(rel_path)
        gui.param1 = True

    def output_select(self, event):
        # When the user selects a directory with the output picker, update the instance variables and display a relative file
        abs_path = self.directory_picker.GetPath()
        gui.save_directory = abs_path
        cwd = os.getcwd()
        rel_path = os.path.relpath(abs_path, cwd)
        self.directory_picker.SetPath(rel_path)
        gui.param2 = True

    def doc_select(self, event):
        # Get the selected document type from the radio box
        input = self.rbox.GetStringSelection()
        templates = {
            'Default': cert_type.DEFAULT.value,
            'SBA': cert_type.SBA.value,
            'DOIU': cert_type.DOIU.value,
            'NOAA': cert_type.NOAA.value
        }
        # Set the global doc_template variable to the selected template file
        gui.doc_template = templates.get(input, cert_type.DEFAULT.value)
        

    def run(self, event):
        # Check if both parameters have been set
        if(gui.param1 and gui.param2):
            test.create_docs(gui.roster_and_grades, gui.save_directory, gui.doc_template)
            # Check for errors
            if(gui.error_1_bool == True):
                gui.error_1_text.Hide()
                gui.error_1_bool = False
            # Display a success message    
            wx.StaticText(self, label = 'Success!', pos=(260,330))
        else:
            # Display an error message
            gui.error_1_text = wx.StaticText(self, label= 'Fill out other fields then use this button!', pos=(260,330))
            gui.error_1_bool = True
        
app = wx.App(0)
gui(None, -1, 'Certificate Creator')
app.MainLoop()