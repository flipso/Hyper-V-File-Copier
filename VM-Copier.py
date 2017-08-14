# That Programm should give you a GUI to Copy Files to an running Hyper-V VM
# Important!: Be sure that the Hyper-V integration services are activated
##You can do this in the VM-Settings under Integration services and you have to check the guest services
# Also you have to change the encoding Variable. If you want to get your active encoding just open up a shell 
##and enter this command: chcp (checkcodepage) 
# The files that you copy to VM arrive everytimes in fp_To = "C:\\Users\\Public\\Documents"

import subprocess, sys, tkinter,os
from tkinter import ttk, filedialog

ENCODING = 'cp850'


def getVms():
	#Open the Powershell and get the available(running) VMS 
    lstVms = []
    cmd = "Get-VM | Where-Object {$_.State -eq 'Running'} | Select name"
    p = subprocess.run(["powershell.exe", cmd], stdout=subprocess.PIPE)  #Get available VMs from powershell

    psRet = p.stdout.decode(ENCODING)
    psRet = psRet.split('\n', 3)[-1] # Remove the first 2 lines
       
    psRet = psRet.splitlines() # Split the lines.(Required to iter over it)
    
    for vmName in psRet: # iter over the lines
        if vmName:
            lstVms.append(vmName.split(' ')[0])   #If name Detected append it and remove spaces
    return lstVms

def getDir():
	#Opens a FileDialog where the user can specify the File, that he wants to Copy
    window.withdraw() #hide Main window

    filePath = filedialog.askopenfilename() #open file dialog
    #fileFolder = filedialog.askdirectory()
    tbFilePath.insert(0, filePath)
    window.deiconify() #show main window

def Copy():
    fp_From = tbFilePath.get().replace('/', '\\')
    fp_To = "C:\\Users\\Public\\Documents"
    vm = cboVms.get()
    cmdFile = "Copy-VMFile \"{}\" -SourcePath \"{}\" -DestinationPath \"{}\" -CreateFullPath -FileSource Host".format(vm, fp_From, fp_To)
        
    p = subprocess.run(["powershell.exe", cmd], stdout=subprocess.PIPE)
    psRet = p.stdout.decode(ENCODING)
    if not psRet:
        print("Files(s) should be transferred")
    else:
        print("ERROR FILE NOT TRANSFERRED")
        if "0x80070050" in psRet:
            print("Sorry, maybe you are an Idiot?! You want to copy a File that already exist!!")
        else:
            print(psRet)


window = tkinter.Tk()
window.title("Hyper-V to VM Copier")
lblVM = ttk.Label(window, text="VM:")
cboVms = ttk.Combobox(window, value=getVms())
lblPath = ttk.Label(window, text="File:")
tbFilePath = ttk.Entry(window)
intFolder = tkinter.IntVar()
rbFolder = ttk.Checkbutton(window, text="Folder", variable=intFolder)
btnFilePath = ttk.Button(window, text="Select", command=getDir)
btnCopy = ttk.Button(window, text="Copy", command=Copy)

lblVM.pack(side=tkinter.LEFT)
cboVms.pack(side=tkinter.LEFT)
lblPath.pack(side=tkinter.LEFT)
tbFilePath.pack(side=tkinter.LEFT)
rbFolder.pack(side=tkinter.LEFT)
btnFilePath.pack(side=tkinter.LEFT) 
btnCopy.pack(side=tkinter.BOTTOM)
window.mainloop()
