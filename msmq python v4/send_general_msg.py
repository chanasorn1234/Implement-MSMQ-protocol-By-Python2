import win32com.client
import os

qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')
qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\myqueue"
queue=qinfo.Open(2,0)   
msg=win32com.client.Dispatch("MSMQ.MSMQMessage")
msg.Label="TestMsg"
msg.Body = "The1"

msg.Send(queue)
queue.Close()

print("send done")
print(computer_name)