from tkinter import *
from tkinter.filedialog import askdirectory, askopenfilename
import stat_mand_modules
window = Tk()
window.title("LearnPro Investigator - NHS GGC")
window.geometry('640x480')


def read_in_file():
    file = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data', title="Choose a learnpro file",
                           )


def get_dir():
    dirname = askdirectory(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                           title="Choose a directory full of learnpro data."
                           )
    dir_label.configure(text=dirname)

def get_indiv_compliance():
    #TODO exception handling
    data = "Input: " + individual_compliance_input.get()
    individual_compliance_label.configure(text=data)

individual_compliance_label = Label(window, text="<text goes here>")
dir_label = Label(window, text="<no directory currently selected.>")
get_dir_button = Button(window, text="Import directory", command=get_dir)
individual_compliance_input = Entry(window, width=50)
individual_compliance_button = Button(window, text="Get user compliance.", command = get_indiv_compliance)
quit_button = Button(window, text="Quit", command = quit)

rad_button_stat_mand = Radiobutton(window, text="Stat mand boxes", value=1, command=stat_mand_modules.read_file)
rad_button_RN_reqs = Radiobutton(window, text="RN Compliance", value=2)
rad_button_manual = Radiobutton(window, text="Select specific modules", value=3)
loading_label = Label(window, text="")


rad_button_stat_mand.pack()
rad_button_RN_reqs.pack()
rad_button_manual.pack()

individual_compliance_label.pack()
individual_compliance_button.pack()
individual_compliance_input.pack()
dir_label.pack()
get_dir_button.pack()
loading_label.pack()
quit_button.pack()
window.mainloop()