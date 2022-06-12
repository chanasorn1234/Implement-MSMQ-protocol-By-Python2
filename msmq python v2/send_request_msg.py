from time import sleep
import win32com.client
import os
import struct
from time import sleep
import pythoncom


destOrig = win32com.client.Dispatch("MSMQ.MSMQDestination")
destResp = win32com.client.Dispatch("MSMQ.MSMQDestination")
msg = win32com.client.Dispatch("MSMQ.MSMQMessage")

res_computer_name = os.getenv('COMPUTERNAME')

destOrig.Formatname = "direct=os:"+res_computer_name+"\\PRIVATE$\\myqueue"
destResp.Formatname = "direct=os:"+res_computer_name+"\\PRIVATE$\\myqueue"

msg.ResponseDestination = destResp

msg.Label = "Test Message: Response"

msg.Send(destOrig)

print("sendrequest done")