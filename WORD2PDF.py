import os
package = "docx2pdf"

try:
    __import__package
except:
    os.system("pip install "+ package)
from docx2pdf import convert

os.system('clear')
os.system('cls')

 #Banner
print("\033[1;35;40m $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ \x1b[0m")
print("\033[0;36;40m                      Khagen Patel                      \x1b[0m")
print("\033[0;36;40m                 WORD to PDF Converter                    \x1b[0m")
print("\033[1;35;40m $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ \x1b[0m")

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
gui = Tk()

gui.geometry("400x125")
gui.title("Khagen Patel")
def getFolderPath():
    folder_selected = filedialog.askdirectory()
    folderPath.set(folder_selected)

def doStuff():
    global folder
    folder = folderPath.get()
    print("Working directory is ", folder)
    gui.destroy()

folderPath = StringVar()
intro= Label (gui, text= "This Script is created by Khagen Patel")
intro.grid(row=1,column = 1)
intro1= Label (gui, text= "Word(.docx) to PDF(.pdf) Converter")
intro1.grid(row=2,column = 1)
a = Label(gui ,text="Select the Folder")
a.grid(row=10,column = 0)
E = Entry(gui,textvariable=folderPath)
E.grid(row=10,column=1)
btnFind = ttk.Button(gui, text="Browse Folder",command=getFolderPath)
btnFind.grid(row=10,column=2)

c = ttk.Button(gui ,text="Convert", command=doStuff)
c.grid(row=15,column=1)
gui.mainloop()



os.chdir(folder)
files = [f for f in os.listdir('.') if os.path.isfile(f)]
for f in files:
	fbasename = os.path.splitext(os.path.basename(f))[0]
	if f.endswith('.docx'):
                print("\nConverting "+f)
                convert(f, os.path.realpath('.') + '/' + fbasename + '.pdf')

print("\nKeep Smiling")
