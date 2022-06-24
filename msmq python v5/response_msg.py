from time import sleep
import win32com.client
import os
import struct
from time import sleep
import pythoncom
import platform
import ConfigParser
import sys
import urllib2

info_SetupEndedReply = None
def receive(pathname):
    qinfo_id=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    keepid = win32com.client.Dispatch("MSMQ.MSMQMessage")
    pre_id = win32com.client.Dispatch("MSMQ.MSMQQueue")

    
    computer_name = os.getenv('COMPUTERNAME')
    pathname = glbconfig_(hostname(),'cc-msqueue')#computer_name+"\\PRIVATE$\\testeroi"
    qinfo_id.FormatName="DIRECT=OS:"+pathname
    pre_id = qinfo_id.Open(1,0)
    keepid = pre_id.Receive()
    # print(keepid)

    ############
    if(keepid.Label == 'SetupTester'):
        res_destination = win32com.client.Dispatch("MSMQ.MSMQDestination")
        resmsg = win32com.client.Dispatch("MSMQ.MSMQMessage")

        ########################## sent to oi
        q_send_oi_1 = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
        computer_name_contract_oi = os.getenv('COMPUTERNAME')
        q_send_oi_1.Formatname = "DIRECT=OS:"+glbconfig_(hostname(),'au-msqueue')#+computer_name_contract_oi+"\\PRIVATE$\\testeroi"
        qsend_oi = q_send_oi_1.Open(2,0)
        msgto_oi = win32com.client.Dispatch("MSMQ.MSMQMessage")
        msgto_oi.Label = keepid.Label
        msgto_oi.Body = keepid.Body
        msgto_oi.Send(qsend_oi)
        qsend_oi.Close()
        print("send step1 to oi done!!")

        while(1):
            q_reciv_oi_1 = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
            q_reciv_q = win32com.client.Dispatch("MSMQ.MSMQQueue")
            q_reciv_m = win32com.client.Dispatch("MSMQ.MSMQMessage")
            q_reciv_oi_1.Formatname = "DIRECT=OS:"+glbconfig_(hostname(),'oi-msqueue')#+computer_name_contract_oi+"\\PRIVATE$\\testeroi"
            q_reciv_q = q_reciv_oi_1.Open(1,0)
            q_reciv_m = q_reciv_q.Peek(pythoncom.Empty, pythoncom.Empty, timeout_sec * 1000)
            if(q_reciv_m != None):
                q_reciv_m = q_reciv_q.Receive()
                print("receive step1 from oi done!!")
                break #q_reciv_m คือสิ่งที่จะเอาไปใช้ต่อ
            q_reciv_q.Close()
            sleep(5)
            
        ##########################
        res_destination = keepid.ResponseDestination
        resmsg.Body = '<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">'+\
            '<Dictionary key="Top">'+\
            '<V dt:dt="i4" key="ReturnCode">0</V>'+\
            '<V dt:dt="string" key="ReturnText">cahansorn</V>'+\
            '<V dt:dt="string" key="ReturnDetails">chanasorn</V>'+\
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
        
        ########################## sent to oi2
        while(1):
            q_reciv_oi_2 = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
            q_reciv_q2 = win32com.client.Dispatch("MSMQ.MSMQQueue")
            q_reciv_m2 = win32com.client.Dispatch("MSMQ.MSMQMessage")
            q_reciv_oi_2.Formatname = "DIRECT=OS:"+glbconfig_(hostname(),'oi-msqueue')#+computer_name_contract_oi+"\\PRIVATE$\\testeroi"
            q_reciv_q2 = q_reciv_oi_2.Open(1,0)
            q_reciv_m2 = q_reciv_q2.Peek(pythoncom.Empty, pythoncom.Empty, timeout_sec * 1000)
            if(q_reciv_m2 != None):
                q_reciv_m2 = q_reciv_q2.Receive()
                break #q_reciv_m2 คือสิ่งที่จะเอาไปใช้ต่อ
            sleep(5)
        # q_send_oi_2 = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
        # q_send_oi_2.Formatname = "DIRECT=OS:"+computer_name_contract_oi+"\\PRIVATE$\\testeroi"
        # qsend_oi2 = q_send_oi_2.Open(2,0)
        # msgto_oi2 = win32com.client.Dispatch("MSMQ.MSMQMessage")
        # msgto_oi2.Label = q_reciv_m.Label
        # msgto_oi2.Body = q_reciv_m2.Body
        # msgto_oi2.Send(qsend_oi2)
        # qsend_oi2.Close()

      

        ##########################
        res_destination = keepid.ResponseDestination
        resmsg.Body = '<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">'+\
            '<Dictionary key="Top">'+\
            '<Dictionary key="BinIndex1">'+\
	        '<V dt:dt="string" key="BinNumber">0</V>'+\
	        '<V dt:dt="string" key="BinDescription">System Error</V>'+\
	        '<V dt:dt="string" key="BinGrade">F</V>'+\
	        '</Dictionary>'+\
            '<Dictionary key="BinIndex2">'+\
	        '<V dt:dt="string" key="BinNumber">2</V>'+\
	        '<V dt:dt="string" key="BinDescription">Pass</V>'+\
	        '<V dt:dt="string" key="BinGrade">P</V>'+\
	        '</Dictionary>'+\
            '<Dictionary key="BinIndex3">'+\
	        '<V dt:dt="string" key="BinNumber">11</V>'+\
	        '<V dt:dt="string" key="BinDescription">OAS</V>'+\
	        '<V dt:dt="string" key="BinGrade">F</V>'+\
	        '</Dictionary>'+\
            '<V dt:dt="string" key="BinCount">38</V>'+\
            '<V dt:dt="string" key="LoadboardID">N/A</V>'+\
            '<V dt:dt="string" key="ContactorID">N/A</V>'+\
            '<V dt:dt="i4" key="ReturnCode">0</V>'+\
            '<V dt:dt="string" key="ReturnText">chanasorn</V>'+\
            '<V dt:dt="string" key="ReturnDetails">chanasorn</V>'+\
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

    if(keepid.Label == 'SetupEndedReply'):
        # info_SetupEndedReply = keepid
        print('receive SetupEndedReply')

    if(keepid.Label == 'EndLot'):
        res_destination = win32com.client.Dispatch("MSMQ.MSMQDestination")
        resmsg = win32com.client.Dispatch("MSMQ.MSMQMessage")

        res_destination = keepid.ResponseDestination
        resmsg.Body = '<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">'+\
            '<Dictionary key="Top">'+\
            '<V dt:dt="i4" key="ReturnCode">0</V>'+\
            '<V dt:dt="string" key="ReturnText"></V>'+\
            '<V dt:dt="string" key="ReturnDetails"></V>'+\
            '<V dt:dt="string" key="SenderID">'+hostname()+'</V>'+\
            '</Dictionary>'+\
            '</Root>'
        resmsg.Label = 'EndLotReply'

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

def get_script_path():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

def hostname():
    #Hostname
    if platform.system() == "Windows":
        hsname = platform.uname()[1]
        hname = hsname.upper()

    else:
        hsname = os.uname()[1]
        hname = str(hsname).upper()
        
    return hname

#config path
# config = ConfigParser.RawConfigParser()
# config.read('%s/config.ini' % get_script_path())

def glbconfig_(section,key):
    # url = config.get('URL','globalConfig-url')
    # f = urllib2.urlopen(url)
    # reply = f.read()
    # f.close()
    glbconfig = ConfigParser.RawConfigParser()
    glbconfig.read('%s/myconf.ini' % get_script_path())
    content = glbconfig.get(section,key)
    return content

while(1):
    pre_qinfo_id = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    pre_keepid = win32com.client.Dispatch("MSMQ.MSMQMessage")
    pre_pre_id = win32com.client.Dispatch("MSMQ.MSMQQueue")

    precheck_computer_name = os.getenv('COMPUTERNAME')
    pre_pathname = glbconfig_(hostname(),'cc-msqueue')#precheck_computer_name+"\\PRIVATE$\\testeroi"
    pre_qinfo_id.FormatName="DIRECT=OS:"+pre_pathname
    pre_pre_id = pre_qinfo_id.Open(1,0)

    timeout_sec = 1.0
    check = pre_pre_id.peek(pythoncom.Empty, pythoncom.Empty, timeout_sec * 1000)
    if(check != None):
        if(check.Label == 'SetupTester'):
            receive('')
        elif(check.Label == 'SetupEndedReply'):
            receive('')
        elif(check.Label == 'EndLot'):
            receive('')
        else:
            sleep(5)
            pre_pre_id.Receive()

    sleep(5)

    
