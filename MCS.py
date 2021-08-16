# Import Modules

from tkinter import *
from openpyxl import *

def submitData():
   
    skuNumber = skuEntry.get()
    serialNumber = serialNumberEntry.get()

    workPerformed = []
    if restoredButtonVar.get():
        workPerformed += ["Restored"]
    if biosButtonVar.get():
        workPerformed += ["BIOS Updated"]
    if cleanedButtonVar.get():
        workPerformed += ["Cleaned Unit"]
    if workPerformedEntry != "":
        workPerformed += [workPerformedEntry.get()]
    workPerformedString = ','.join(map(str,workPerformed))

    otherRemarks = []
    if adaptorButtonVar.get():
        otherRemarks += ["Missing Adaptor"]
    if boxButtonVar.get():
        otherRemarks += ["Unit is in Wrong Box"]
    if damagedButtonVar.get():
        otherRemarks += ["Damaged Unit"]
    if remarksEntry != None:
        otherRemarks += [remarksEntry.get()]
    otherRemarksString = ','.join(map(str,otherRemarks))

    data = (skuNumber,serialNumber,workPerformedString,otherRemarksString)
    print(data)

    #wbName = 'computer.xlsx'
    #wb = load_workbook(wbName)
    #page=wb.active
    #data = [()]
    #page.append(data)
    #wb.save(filename=wbName)

window = Tk()
window.title('Markan Computer Services')
w,h = window.winfo_screenwidth(), window.winfo_screenheight()
window.geometry("%dx%d+0+0" % (w,h))

# SKU

skuLabel = Label(window,text="SKU",font=('Helvetica 25 bold'))
skuLabel.config(anchor=CENTER)
skuLabel.pack()
skuEntry1 = Entry(window,justify="center")
skuEntry1.pack()

# S/N

serialNumberLabel = Label(window,text="S/N",font=('Helvetica 25 bold'))
serialNumberLabel.config(anchor=CENTER)
serialNumberLabel.pack()
serialNumberEntry = Entry(window,justify="center")
serialNumberEntry.pack()

# Work Performed
workPerformedLabel = Label(window,text="Work Performed",font=('Helvetica 25 bold'))
workPerformedLabel.config(anchor=CENTER)
workPerformedLabel.pack()

restoredButtonVar = IntVar()
restoredButton = Checkbutton(window,text="Restored",variable = restoredButtonVar)
restoredButton.config(anchor=CENTER)
restoredButton.pack()

biosButtonVar = IntVar()
biosButton = Checkbutton(window,text="BIOS Updated",variable = biosButtonVar)
biosButton.config(anchor=CENTER)
biosButton.pack()

cleanedButtonVar = IntVar()
cleanedButton = Checkbutton(window,text="Cleaned",variable = cleanedButtonVar)
cleanedButton.config(anchor=CENTER)
cleanedButton.pack()

workPerformedOtherLabel = Label(window,text="Other tasks",font=('Helvetica 15 bold'))
workPerformedOtherLabel.config(anchor=CENTER)
workPerformedOtherLabel.pack()

workPerformedEntry = Entry(window,justify="center")
workPerformedEntry.pack()

# Other Remarks
remarksLabel = Label(window,text="Other Remarks",font=('Helvetica 25 bold'))
remarksLabel.config(anchor=CENTER)
remarksLabel.pack()

adaptorButtonVar = IntVar()
adaptorButton = Checkbutton(window,text="Missing Adaptor",variable = adaptorButtonVar)
adaptorButton.config(anchor=CENTER)
adaptorButton.pack()

boxButtonVar = IntVar()
boxButton = Checkbutton(window,text="Unit is in Wrong Box",variable = boxButtonVar)
boxButton.config(anchor=CENTER)
boxButton.pack()

damagedButtonVar = IntVar()
damagedButton = Checkbutton(window,text="Damaged Unit",variable = damagedButtonVar)
damagedButton.config(anchor=CENTER)
damagedButton.pack()

remarksOtherLabel = Label(window,text="Other Remarks",font=('Helvetica 15 bold'))
remarksOtherLabel.config(anchor=CENTER)
remarksOtherLabel.pack()

remarksEntry = Entry(window,justify="center")
remarksEntry.pack()

# Submit
submitButton = Button(window, text="Submit",command=submitData)
submitButton.config(anchor=CENTER)
submitButton.pack()

window.mainloop()
