# from unittest import result
import win32com.client
import os
import re
import struct
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
    
queue.Close()