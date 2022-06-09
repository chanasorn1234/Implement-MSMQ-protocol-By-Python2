# from unittest import result
from time import sleep
import win32com.client
import os
import re
import struct
from time import sleep
import pythoncom
pattern = re.compile('x')
qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')
qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\myqueue"
queue=qinfo.Open(1,0)   # Open a ref to queue to read(1)
msg=queue.Receive()
print("Label:",msg.Label)
print("Body :",msg.Body)
print("ID:",str(msg.SourceMachineGuid))
print("ID2:",msg.LookupId)
print("Time:",msg.SentTime)
# print(int(bytes(msg.Id).e),"")
result = msg.Id[16:20]
# result = b'\x0f\xc8\x01\x00'
print(type(result))
print(result,'')
print(struct.unpack('4B',result))
num = struct.unpack('4B',result)
message_num = 0
for i in num:
    message_num += i
print(message_num)
queue.Close()

sleep(1)

resqinfo = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
res_computer_name = os.getenv('COMPUTERNAME')
resqinfo.FormatName="direct=os:"+res_computer_name+"\\PRIVATE$\\myqueue"
resqeue = resqinfo.Open(2,0)
resmsg = win32com.client.Dispatch("MSMQ.MSMQMessage")
resmsg.Body = "KuyKob"
resmsg.Label = "ResTestMsg"
print(type(msg.Id))
resmsg.CorrelationId = msg.Id
res_destination = win32com.client.Dispatch("MSMQ.MSMQDestination")
res_destination.FormatName = "direct=os:"+computer_name+"\\PRIVATE$\\myqueue"
resmsg.ResponseDestination = res_destination
resmsg.Send(resqeue)
resqeue.Close()






















# result = int(result.encode('hex'), 16)
# print(result,"")

# result = re.sub(pattern,'',result)
# print(result)
# result = str(msg.Id[16:])
# print(msg.Id)
# msg2=queue.ReceiveById()
# print(msg2)
# try:
#     print("ID3:",int(result[16:18],16))
# except:
#     print("ID3:",int(result[17],16))