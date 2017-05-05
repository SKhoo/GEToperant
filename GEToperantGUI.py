from tkinter import *
from tkinter.filedialog import *
from tkinter.messagebox import *
##from tkinter.filedialog import askopenfilenames
##from tkinter.filedialog import asksaveasfilename
import GEToperant

GETprofile = None
MPC_filenames = None

#Define functions for the menus
def openprofile():
    global GETprofile
    GETprofile = askopenfilename(title = 'Select data profile', filetypes =  [('Excel', '*.xlsx')])
def opendata():
    global MPC_filenames
    MPC_filenames = askopenfilenames(title = 'Select files to import')
def saveoutput():
    outputfile = asksaveasfilename(title = 'Save output file as', defaultextension='.xlsx', filetypes=(('Excel', '*.xlsx'),('All Files', '*.*')))
    if GETprofile == None or MPC_filenames == None:
        showerror('Error! Profile or data file selection.', 'Please select a data profile and Med-PC data files first.')
    else:
        GEToperant.GEToperant(GETprofile, MPC_filenames, outputfile)

def GETexpress():
    GETprofile = askopenfilename(title = 'Select data profile', filetypes =  [('Excel', '*.xlsx')])
    MPC_filenames = askopenfilenames(title = 'Select files to import')
    outputfile = asksaveasfilename(title = 'Save output file as', defaultextension='.xlsx', filetypes=(('Excel', '*.xlsx'),('All Files', '*.*')))
    GEToperant.GEToperant(GETprofile, MPC_filenames, outputfile)

def helpme():
    helpwindow = Toplevel()
    helpwindow.title('How to use GEToperant')
    helptext = Text(helpwindow, height = 30, width = 80)
    helptext.pack(side= 'top')
    scroll = Scrollbar(helpwindow, command = helptext.yview)
    helptext.configure(yscrollcommand = scroll.set)
    helptext.tag_configure('regular', font=('Verdana', 12))
    howtoGET = """
    How to use GEToperant

    Using GEToperant involves three steps.
    1. Select a data profile
    2. Open your Med PC raw data file(s)
    3. Save your output

    Your data profile tells GEToperant what data you want extracted
    and what to label each element as. You can extract:
    - a single element
    - a section of an array
    - a whole array

    Your data profile needs to have up to 7 pieces of information:
    1. A Label
    2. A Label Start Value
    3. A Label Increment
    4. An Array or Variable
    5. The Start Element
    6. The Increment Element
    7. The Stop Element

    In order to extract a single element you will need to define:
    - The Label
    - The Array or Variable
    - The Start Element (i.e. the element you want extracted)
    - The Increment Element (which must equal 0)

    For example, the element A(0) contains the total lever responses.
    You would define the label as 'Lever Presses', the Array as 'A',
    the Start Element as 0 and the Increment Element as 0. This tells
    GEToperant to get the element A(0) from all sessions in the data
    files you load and to label it 'Lever Presses'.

    In order to extract a section of an array you need:
    - The label
    - The Array or Variable
    - The Start Element
    - The Increment Element
    - The Stop Element
    You can also use:
    - The Label Start Value
    - The Label Increment

    Your Stop Element must be greater than your Start Element and
    your Increment Element must be greater than 0. This will tell
    GEToperant to start at a particular part of the array and keep
    going up by the increments you define until it reaches the Stop
    Element. So if you wanted every second value of the B array from
    beginning to element 30, you would set the Start Element to 0,
    the Incremenet Element to 2 and the Stop Element to 30.

    The Label Increment and Label Start Value are optional and allow
    you to define a value to put at the end of your label. This is
    useful for a series like timebins. For example, you could have
    a label of 'Responses Min' with a Label Start Value of 1 and a
    Label Increment of 1. You would then get 'Responses Min 1',
    'Responses Min 2', 'Responses Min 3' and so on.

    In order to extract an array until it ends you will need the same
    details as required to extract a section of an array except you
    should leave the Stop Element blank or write something in it, such
    as 'End'. However, any text string will be read as the end of the
    array.

    The session headers are extracted automatically, but session comments
    are not extracted automatically. In order to extract comments
    provide:
    - The Label
    - An Array or Variable with the word 'comment' in it (this is not
    case sensitive)
    - A Start Element and Increment Element of 0

    Once you have your data profile, all you need to do is press the buttons!
    """
    helptext.insert(END, howtoGET, 'regular')
    helptext.pack(side=LEFT)
    scroll.pack(side=RIGHT, fill = Y)

def aboutGET():
    aboutme = Toplevel()
    aboutme.title('About GEToperant')
    abouttext = Text(aboutme, height = 22, width = 75)
    abouttext.pack(side= 'top')
    abouttext.tag_configure('regular', font=('Verdana', 12))
    about = """
    GEToperant is a general extraction tool for Med-PC®.
    It was designed to be compatible with Med-PC® IV but given how
    little Med PC changes, it should be compatible with Med-PC® V.
    It was written by Shaun Khoo (ORCID: 0000-0002-0972-3788).
    It is free open source software available under an MIT license.
    You pay nothing and you can do with it as you please.
    The MIT license allows you to sell GEToperant but I have no
    idea how you can make money given that it is available for free.

    If you have enjoyed using GEToperant, please reference it
    in one of your publications. You can refer to the url or cite
    one of my papers.

    For more information check it out on Github:
    www.github.com/Skhoo
    For up to date contact information visit:
    https://orcid.org/0000-0002-0972-3788
    """
    abouttext.insert(END, about, 'regular')
    abouttext.pack(side=LEFT)

def licenseMIT():
    licenseme = Toplevel()
    licenseme.title('About GEToperant')
    MIT = Text(licenseme, height = 31, width = 75)
    MIT.pack(side= 'top')
    MIT.tag_configure('regular', font=('Verdana', 12))
    MITtext = """
    MIT License

    Copyright (c) 2017 Shaun Khoo

    Permission is hereby granted, free of charge, to any person
    obtaining a copy of this software and associated documentation
    files (the "Software"), to deal in the Software without restriction,
    including without limitation the rights to use, copy, modify, merge,
    publish, distribute, sublicense, and/or sell copies of the Software,
    and to permit persons to whom the Software is furnished to do so,
    subject to the following conditions:

    The above copyright notice and this permission notice shall be
    included in all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY
    OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
    NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND
    NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
    DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
    CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF
    OR IN CONNECTION WITH THE SOFTWARE OR THE USE
    OR OTHER DEALINGS IN THE SOFTWARE.
    """
    MIT.insert(END, MITtext, 'regular')
    MIT.pack(side=LEFT)

def quit():
    root.destroy()

root = Tk()

##Set window size
root.geometry('876x420')
root.title('GEToperant v0.9a >(\' . \')<')
Montre = PhotoImage(file='icon.pnm')
root.wm_iconphoto('True', Montre)

#Display header logo
Kip = PhotoImage(file='logo.pnm')
displaylogo = Label(root, image = Kip).pack(side = 'top')

## Menus
menu = Menu(root)
root.config(menu = menu)
filemenu = Menu(menu)
menu.add_cascade(label = 'File', menu = filemenu)
filemenu.add_command(label = 'Select Profile', command = openprofile)
filemenu.add_command(label = 'Open Data File(s)', command = opendata)
filemenu.add_command(label = 'Save Output As', command = saveoutput)
filemenu.add_separator()
filemenu.add_command(label = 'Close', command = root.quit)
                                 
helpmenu = Menu(menu)
menu.add_cascade(label = 'Help', menu = helpmenu)
helpmenu.add_command(label = 'How to use GEToperant', command = helpme)
helpmenu.add_command(label = 'About', command = aboutGET)
helpmenu.add_command(label = 'License', command = licenseMIT)

## Buttons
class App:
    def __init__(self, master):
        frame = Frame(master)
        frame.pack()
        self.express = Button(frame, text = 'GEToperant Express', command = GETexpress)
        self.express.pack(padx = 10, pady = 50, side=LEFT)
        self.profile = Button(frame, text = 'Select Profile', command = openprofile)
        self.profile.pack(padx = 10, pady = 50, side=LEFT)
        self.MPCdatafiles = Button(frame, text = 'Select Med-PC data file(s)', command = opendata)
        self.MPCdatafiles.pack(padx = 10, pady = 50, side=LEFT)
        self.execall = Button(frame, text = 'Select save file data', command = saveoutput)
        self.execall.pack(padx = 10, pady = 50, side=LEFT)
        self.exit = Button(frame, text = 'Quit', command = quit)
        self.exit.pack(padx = 10, pady = 50, side=LEFT)

app = App(root)
root.mainloop()
