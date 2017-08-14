# That Programm should give you a GUI to Copy Files to an running Hyper-V VM
# Important!: Be sure that the Hyper-V integration services are activated
##You can do this in the VM-Settings under Integration services and you have to check the guest services
# Also you have to change the encoding Variable. If you want to get your active encoding just open up a shell 
##and enter this command: chcp (checkcodepage) 
# The files that you copy to VM arrive everytimes in DEST_PATH = "C:\\Users\\Public\\Documents"

import subprocess, sys, tkinter,os
from tkinter import ttk, filedialog

ENCODING = "cp850"
DEST_PATH = "C:\\Users\\Public\\Documents"

def getVms():
    #Open the Powershell and get the available(running) VMS 
    lstVms = []

    cmd = "Get-VM | Where-Object {$_.State -eq 'Running'} | Select name"
    p = subprocess.run(["powershell.exe", cmd], stdout=subprocess.PIPE)  #Get available VMs from powershell

    sPsRet = p.stdout.decode(ENCODING)
    sPsRet = sPsRet.split('\n', 3)[-1] # Remove the first 2 lines
    sPsRet = sPsRet.splitlines() # Split the lines.(Required to iter over it)
    
    for sVmName in sPsRet: # iter over the lines
        if sVmName:
            lstVms.append(sVmName.split(' ')[0])   #If name Detected append it and remove spaces at the end of string

    if not lstVms:  # If no vm detected wit state 'Running' then exit
        print("Iam Sorry, i have to exit!")
        print("You have to start a Hyper-V Vm first to copy something!")
        sys.exit()
    return lstVms

def getUserDir():
    #Opens a FileDialog where the user can specify the File or Fodlder

    window.withdraw() #hide Main window
    if intFolder.get() == 1:
        sPath = filedialog.askdirectory() 
    else:
        sPath = filedialog.askopenfilename() #open file dialog
    
    tbFilePath.insert(0, sPath) #fill textbox with path
    window.deiconify() #show main window

def Copy():
    #Prepare the path string and runs Powershell copy-routine

    fp_From = tbFilePath.get().replace('/', '\\')
    sVM = cboVms.get()

    sCmdFile = "Copy-VMFile \"{}\" -SourcePath \"{}\" -DestinationPath \"{}\" -CreateFullPath -FileSource Host".format(sVM, fp_From, DEST_PATH)
    sCmdFolder = "Get-ChildItem \"{}\"  | % {{ Copy-VMFile \"{}\"  -SourcePath $_.FullName -DestinationPath \"{}\"  -CreateFullPath -FileSource Host}}".format(fp_From, sVM, DEST_PATH)
    
    if intFolder.get() == 1:
        cmd = sCmdFolder
    else:
        cmd = sCmdFile
   
    p = subprocess.run(["powershell.exe", cmd], stdout=subprocess.PIPE)
    sPsRet = p.stdout.decode(ENCODING)

    if not sPsRet:
        print("Files(s) should be transferred")
    else:
        print("ERROR FILE NOT TRANSFERRED")
        if "0x80070050" in sPsRet:
            print("Sorry, maybe you are an Idiot?! You want to copy a File that already exist!!")
        else:
            print(sPsRet)



#Controls
window = tkinter.Tk()
window.title("Hyper-V to VM Copier")
lblVM = ttk.Label(window, text="VM:")
cboVms = ttk.Combobox(window, value=getVms())
lblPath = ttk.Label(window, text="File/Folder:")
tbFilePath = ttk.Entry(window, width=60)
intFolder = tkinter.IntVar()
rbFolder = ttk.Checkbutton(window, text="Folder", variable=intFolder)
btnFilePath = ttk.Button(window, text="Select", command=getUserDir)
btnCopy = ttk.Button(window, text="Copy", command=Copy)

lblVM.pack(side=tkinter.LEFT)
cboVms.pack(side=tkinter.LEFT)
lblPath.pack(side=tkinter.LEFT)
tbFilePath.pack(side=tkinter.LEFT)
rbFolder.pack(side=tkinter.LEFT)
btnFilePath.pack(side=tkinter.LEFT) 
btnCopy.pack()

window.mainloop() 
