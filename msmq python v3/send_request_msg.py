
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

    pathname = res_computer_name+"\\PRIVATE$\\myqueue" #"mth-vm-b67871"
    pathname2 = "mth-vm-b67871"+"\\PRIVATE$\\testeroi"
    destOrig.Formatname = "DIRECT=OS:"+ pathname
    destResp.Formatname = "DIRECT=OS:"+ pathname

    msg.ResponseDestination = destResp
    ##########################
    msg.Body = str(msgs)
    msg.Label = str(lmsgs)
    if(msg.Label == "SetupEnded"):
        print("Send Correlation ID: ", str(mId),type(mId))
    
    elif(msg.Label == "SetupTester"):
        # l3.config(text = 'Status : Start Lot')
        # l3.config(text = 'Status : Recieved [%s]' % str(lmsgs))
        msg.Body += '<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">'+\
        '<Dictionary key="Top">'+\
        '<V dt:dt="string" key="ProductID">LEAK1TN2XAXF</V>'+\
        '<V dt:dt="string" key="LotID">MMT-230401629.000</V>'+\
        '<V dt:dt="string" key="HandlerID">TapestryZ</V>'+\
        '<V dt:dt="i2" key="TempSetpoint">25</V>'+\
        '<V dt:dt="string" key="OperatorID">finalop</V>'+\
        '<V dt:dt="string" key="PartNum">24F16KA102</V>'+\
        '<V dt:dt="string" key="DeviceType">OTP</V>'+\
        '<V dt:dt="string" key="CPOnChecksum">0</V>'+\
        '<V dt:dt="string" key="CPOffChecksum">0</V>'+\
        '<V dt:dt="string" key="QCode">0</V>'+\
        '<V dt:dt="string" key="TestProgFileName">P:\LEAK0\LEAK0_B56\LEAK0_B56.XLS</V>'+\
        '<V dt:dt="string" key="TestProgChecksum">29187440</V>'+\
        '<V dt:dt="string" key="HardwareMap">x24mct28ssop</V>'+\
        '<V dt:dt="string" key="TestFlow">f1-prd-std-28L</V>'+\
        '<V dt:dt="string" key="Environment">FS1@25C</V>'+\
        '<V dt:dt="string" key="HandlerType">Tapestry PH-1</V>'+\
        '<V dt:dt="string" key="SenderID">TapestryZ</V>'+\
        '<V dt:dt="string" key="TestMode">FT</V>'+\
        '</Dictionary>'+\
        '</Root>'

    ##########################
    msg.Send(destOrig)

    print("sendrequest done")

# def callws():
#     pass

# def glbconfig_():
#     pass

send(1,'SetupTester','','')
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