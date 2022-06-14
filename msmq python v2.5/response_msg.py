from time import sleep
import win32com.client
import os
import struct
from time import sleep
import pythoncom
import platform

def receive(pathname):
    qinfo_id=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    keepid = win32com.client.Dispatch("MSMQ.MSMQMessage")
    pre_id = win32com.client.Dispatch("MSMQ.MSMQQueue")


    computer_name = os.getenv('COMPUTERNAME')
    pathname = computer_name+"\\PRIVATE$\\myqueue"
    qinfo_id.FormatName="DIRECT=OS:"+pathname
    pre_id = qinfo_id.Open(1,0)
    timeout_sec = 1.0
    keepid = pre_id.Receive()
    # print(keepid)

    ############
    if(keepid.Label == 'SetupTester'):
        res_destination = win32com.client.Dispatch("MSMQ.MSMQDestination")
        resmsg = win32com.client.Dispatch("MSMQ.MSMQMessage")
        res_computer_name = os.getenv('COMPUTERNAME')

        res_destination = keepid.ResponseDestination
        resmsg.Body = '<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">'+\
            '<Dictionary key="Top">'+\
            '<V dt:dt="i4" key="ReturnCode">0</V>'+\
            '<V dt:dt="string" key="ReturnText"></V>'+\
            '<V dt:dt="string" key="ReturnDetails"></V>'+\
            '<V dt:dt="string" key="SenderID">'+hostname()+'</V>'+\
            '</Dictionary>'+\
            '</Root>'
        resmsg.Label = 'SetupTesterReply'

        resmsg.CorrelationId = keepid.Id
        
        print(type(keepid.Id))
        print(struct.unpack("<HH",resmsg.CorrelationId[16:20]),struct.unpack("<HH",keepid.Id[16:20]))
        
        resmsg.Send(res_destination)
        res_destination.Close()
        pre_id.Close()

    if(keepid.Label == 'SetupTester'):
        res_destination = win32com.client.Dispatch("MSMQ.MSMQDestination")
        resmsg = win32com.client.Dispatch("MSMQ.MSMQMessage")
        res_computer_name = os.getenv('COMPUTERNAME')

        res_destination = keepid.ResponseDestination
        resmsg.Body = '<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">'+\
            '<Dictionary key="Top">'+\
            '<V dt:dt="i4" key="ReturnCode">-1</V>'+\
            '<V dt:dt="string" key="ReturnText">Program Not Loaded</V>'+\
            '<V dt:dt="string" key="ReturnDetails"></V>'+\
            '<V dt:dt="string" key="SetupTesterMsgID">{CDC5E747-EE4F-4A9A-A8E4-7F7F8709FAC4}\\23031291</V>'+\
            '<V dt:dt="string" key="SenderID">'+hostname()+'</V>'+\
            '</Dictionary>'+\
            '</Root>'
        resmsg.Label = 'SetupEnded'

        resmsg.CorrelationId = keepid.Id
        
        print(type(keepid.Id))
        print(struct.unpack("<HH",resmsg.CorrelationId[16:20]),struct.unpack("<HH",keepid.Id[16:20]))
        
        resmsg.Send(res_destination)
        res_destination.Close()
        pre_id.Close()

    ##############
    if(res_destination == None):
        res_destination = win32com.client.Dispatch("MSMQ.MSMQDestination")
        resmsg = win32com.client.Dispatch("MSMQ.MSMQMessage")
        res_computer_name = os.getenv('COMPUTERNAME')

        res_destination = keepid.ResponseDestination
        resmsg.Body = "KhunKob"
        resmsg.Label = "ResTestMsg"

        resmsg.CorrelationId = keepid.Id
        # print(keepid.Id[16:20].tobytes())
        print(type(keepid.Id))
        # print(keepid.Destination)
        print(struct.unpack("<HH",resmsg.CorrelationId[16:20]),struct.unpack("<HH",keepid.Id[16:20]))


        resmsg.Send(res_destination)
        res_destination.Close()
        pre_id.Close()

def hostname():
    #Hostname
    if platform.system() == "Windows":
        hsname = platform.uname()[1]
        hname = hsname.upper()

    else:
        hsname = os.uname()[1]
        hname = str(hsname).upper()
        
    return hname

receive('')