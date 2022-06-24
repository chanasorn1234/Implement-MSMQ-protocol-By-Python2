
import win32com.client
import os
import struct
from time import sleep
import pythoncom
# import Tkinter as tk 
# import tkMessageBox 
def send(mId,lmsgs,msgs,pathname2):
    destOrig = win32com.client.Dispatch("MSMQ.MSMQDestination")
    destResp = win32com.client.Dispatch("MSMQ.MSMQDestination")
    msg = win32com.client.Dispatch("MSMQ.MSMQMessage")

    res_computer_name = os.getenv('COMPUTERNAME')

    pathname = res_computer_name+"\\PRIVATE$\\testeroi" #"mth-vm-b67871"
    pathname2 = "ss-srv5"+"\\PRIVATE$\\myqueue"
    destOrig.Formatname = "DIRECT=OS:"+ pathname2
    destResp.Formatname = "DIRECT=OS:"+ pathname

    msg.ResponseDestination = destResp
    ##########################
    msg.Body = str(msgs)
    msg.Label = str(lmsgs)
    if(msg.Label == "SetupEnded"):
        print("Send Correlation ID: ", str(mId),type(mId))
    
    elif(msg.Label == "smGetTestBinGroups"):
        # l3.config(text = 'Status : Start Lot')
        # l3.config(text = 'Status : Recieved [%s]' % str(lmsgs))
        msg.Body += '<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">'+\
        '<Dictionary key="Top">'+\
        '<V dt:dt="string" key="TestBinGroup">LEBC0</V>'+\
        '</Dictionary>'+\
        '</Root>'

    ##########################
    msg.Send(destOrig)

    print("sendrequest done")

# def callws():
#     pass

# def glbconfig_():
#     pass

send(1,'smGetTestBinGroups','','')
# app = tk.Tk()
# app.title('Autoload tool')
# app.geometry("440x130")
# app.eval('tk::PlaceWindow . center')

# l1 = tk.Label(app,font =("Courier", 12))
# l2 = tk.Label(app,font =("Courier", 12))
# l3 = tk.Label(app,font =("Courier", 12))

# l1.pack(padx= (30,0),pady=(15,0),anchor ='w')
# l2.pack(padx= (30,0),pady=(5,0),anchor ='w')
# l3.pack(padx= (30,0),pady=(5,0),anchor ='w')

# app.mainloop()