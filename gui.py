import wx
import test
from cert_type_enum import cert_type

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
        wx.Dialog.__init__(self, parent, id, title,
                           size=(APP_SIZE_X, APP_SIZE_Y))                   

        # Add buttons and radio box to the dialog
        radio_list = ['Default', 'SBA', 'DOIU','NOAA']
        wx.Button(self, 1, 'Select Roster and Grades', (30, 50))
        wx.Button(self, 2, 'Output Directory', (30, 90))
        wx.Button(self, 3, 'Create Certificates', (350, 300))
        self.rbox = wx.RadioBox(self, 4, 'Certificate Type', choices=radio_list, pos=(30, 130))

        # Bind events to the buttons and radio box
        self.Bind(wx.EVT_BUTTON, self.file_select, id=1)
        self.Bind(wx.EVT_BUTTON, self.output_select, id=2)
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
        gui.roster_and_grades = test.select_wb()
        wx.StaticText(self, label=gui.shorten(gui.roster_and_grades), pos=(190, 50))
        gui.param1 = True

    def output_select(self, event):
        gui.save_directory = test.select_save()
        wx.StaticText(self, label=gui.shorten(gui.save_directory), pos=(150,90))
        gui.param2 = True

    def doc_select(self,event):
        input = self.rbox.GetStringSelection()
        match input:
            case 'Default':
                gui.doc_template = cert_type.DEFAULT.value
            case 'SBA':
                gui.doc_template = cert_type.SBA.value
            case 'DOIU':
                gui.doc_template = cert_type.DOIU.value
            case 'NOAA':
                gui.doc_template = cert_type.NOAA.value
            case _:
                gui.doc_template = cert_type.DEFAULT.value
     

    def run(self, event):
        if(gui.param1 and gui.param2):
            test.create_docs(gui.roster_and_grades, gui.save_directory, gui.doc_template)
            if(gui.error_1_bool == True):
                gui.error_1_text.Hide()
                gui.error_1_bool = False
            wx.StaticText(self, label = 'Success!')
        else:
            gui.error_1_text = wx.StaticText(self, label= 'Fill out other fields then use this button!')
            gui.error_1_bool = True
            

app = wx.App(0)
gui(None, -1, 'Certificate Creator')
app.MainLoop()