from time import sleep
import win32com.client
import os
import struct
from time import sleep
import pythoncom


qinfo_id=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
keepid = win32com.client.Dispatch("MSMQ.MSMQMessage")
pre_id = win32com.client.Dispatch("MSMQ.MSMQQueue")


computer_name = os.getenv('COMPUTERNAME')
qinfo_id.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\myqueue"
pre_id = qinfo_id.Open(1,0)
timeout_sec = 1.0
keepid = pre_id.Receive()
# print(type(keepid.Id[16:20]))
# print(pre_id)


res_destination = win32com.client.Dispatch("MSMQ.MSMQDestination")
resmsg = win32com.client.Dispatch("MSMQ.MSMQMessage")
res_computer_name = os.getenv('COMPUTERNAME')

res_destination = keepid.ResponseDestination

if(res_destination != None):

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

else:
    print("Request message type invalid")














# qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
# computer_name = os.getenv('COMPUTERNAME')
# qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\myqueue"
# queue=qinfo.Open(1,0)   # Open a ref to queue to read(1)
# msg=queue.Receive()
# print("Label:",msg.Label)
# print("Body :",msg.Body)
# print("ID:",str(msg.SourceMachineGuid))
# print("ID2:",msg.LookupId)
# print("Time:",msg.SentTime)
# print(type(msg.Id))
# result = msg.Id[16:20]
# result2 = msg.SourceMachineGuid
# # result = b'\x0f\xc8\x01\x00'

# print(result,'')
# print(result2,'')
# result2 = result2.encode('utf_8')
# print(result2)
# num = struct.unpack('<HH',result)
# message_num = 0
# for i in num:
#     message_num += i
# print(message_num)
# queue.Close()

# # result = bytes(result)
# # print(result)
# # print(type(result))
# # sleep(1)

# frame = bytearray()
# for i in range(0,20):
#     frame.append(msg.Id[i])
# print(frame)
# print(len(frame))


# frame = bytearray()
# for i in range(0,20):
#     frame.append(keepid.Id[i])

# print(frame)