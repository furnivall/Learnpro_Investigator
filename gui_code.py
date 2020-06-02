from tkinter import *
from tkinter.filedialog import askdirectory, askopenfilename
# import stat_mand_modules
import pandas as pd

window = Tk()
window.title("LearnPro Investigator - NHS GGC")
window.geometry('640x480')
file = pd.DataFrame()


def read_in_file():
    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                               title="Choose a learnpro file")
    global file
    file = pd.read_csv(filename, sep="\t", skiprows=14)
    loading_label.configure(text=file.columns)


def get_dir():
    dirname = askdirectory(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                           title="Choose a directory full of learnpro data."
                           )
    dir_label.configure(text=dirname)


def get_indiv_compliance():
    # TODO exception handling
    data = "Input: " + individual_compliance_input.get()
    individual_compliance_label.configure(text=data)


individual_compliance_label = Label(window, text="<text goes here>")
dir_label = Label(window, text="<no directory currently selected.>")
get_dir_button = Button(window, text="Import directory", command=get_dir)
individual_compliance_input = Entry(window, width=50)
individual_compliance_button = Button(window, text="Get user compliance.", command=get_indiv_compliance)
quit_button = Button(window, text="Quit", command=quit)

read_file = Button(window, text="Read in file", command=read_in_file)

loading_label = Label(window, text="No file currently imported.")

read_file.pack(side=TOP)
get_dir_button.pack(side=TOP)
individual_compliance_label.pack(side=LEFT)
individual_compliance_button.pack(side=LEFT)
individual_compliance_input.pack()
dir_label.pack()
get_dir_button.pack()
loading_label.pack()
quit_button.pack()
window.mainloop()
