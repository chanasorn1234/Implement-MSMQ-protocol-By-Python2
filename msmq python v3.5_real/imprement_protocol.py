from __future__ import unicode_literals
from base64 import encode
import unicodedata
import httplib
from telnetlib import theNULL
import threading 
import win32com.client
from win32com.client import gencache
import xml.etree.ElementTree as ET
import urllib
import urllib2 
import Tkinter as tk 
import ConfigParser
import json
import pythoncom
import tkMessageBox 
import os, platform
import logging
from logging.handlers import TimedRotatingFileHandler
import types
import sys
from datetime import date
import binascii
import time

today = date.today()

#script_dir = os.getcwd()
def get_script_path():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

#Hostname
if platform.system() == "Windows":
    hsname = platform.uname()[1]
    hname = hsname.upper()
    print('Host : %s' % hname)
else:
    hsname = os.uname()[1]
    hname = str(hsname).upper()
    print('Host : %s' % hname)

#config path
config = ConfigParser.RawConfigParser()
config.read('%s/config.ini' % get_script_path())

def glbconfig_(section,key):
    url = config.get('URL','globalConfig-url')
    f = urllib2.urlopen(url)
    reply = f.read()
    f.close()

    with open('%s/myconf.ini' % get_script_path(), "wb") as myfile:
        myfile.write(reply)

    glbconfig = ConfigParser.RawConfigParser()
    glbconfig.read('%s/myconf.ini' % get_script_path())
    content = glbconfig.get(section,key)
    return content

#Log file
if not os.path.exists('%s/log' % get_script_path()) :
    os.makedirs('%s/log' % get_script_path())
if not os.path.exists('%s/log/%s' % (get_script_path(),today.strftime("%Y-%m"))) :
    os.makedirs('%s/log/%s' % (get_script_path(),today.strftime("%Y-%m")))

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter(fmt=' %(asctime)s   [%(levelname)s]   |  %(message)s',
                                        datefmt='%m-%d-%y %H:%M:%S')
fh = TimedRotatingFileHandler('%s/log/%s/log.log' % (get_script_path(),today.strftime("%Y-%m")), when="midnight", interval=1, encoding='utf8')
fh.setFormatter(formatter)
logger.addHandler(fh)

import array
import struct
def receive(pathname):
    #try:
        qinfo = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
        qinfo.FormatName = 'direct=os:'+pathname 
      
        try:
            myq = qinfo.Open(1,0)

        except:
            from win32com.client import gencache
            msmq = gencache.EnsureModule('{D7D6E071-DCCD-11D0-AA4B-0060970DEBAE}', 0, 1, 0)
            qinfo = msmq.MSMQQueueInfo()
            qinfo.PathName = pathname 
            qinfo.Create()
            logger.info('Create message queue: %s ' % qinfo.PathName) 
            myq = qinfo.Open(1,0)

        msg = win32com.client.Dispatch("MSMQ.MSMQMessage")
        if myq.IsOpen:
            timeout_sec = 1.0
            if myq.Peek(pythoncom.Empty, pythoncom.Empty, timeout_sec * 1000):
                msg = myq.Receive()
                if msg is not None:
                    by = msg.Id[16:20] #b'\x0f\xc8\x01\x00'
                    '''print(bytes(msg.Id),bytes(by))
                    print(struct.unpack('<HH', by))
                    #print(int.from_bytes(by,'big')) error int has no attribute from_bytes
                    binlst = [bin(ord(c))[2:].rjust(8,str('0')) for c in bytes(by)]  # remove '0b' from string, fill 8 bits
                    binstr = ''.join(binlst)
                    print(binlst,binstr,int(binstr, 2))
                    for i in binlst:
                        print(i,int(str(i), 2))
                    b = binascii.hexlify(msg.Id)
                    print ("GUID: %s" % str(msg.SourceMachineGuid))
                    print ("ID: ", msg.Id, binascii.hexlify(msg.Id).upper(),binascii.hexlify(by).upper())
                    print ("CID: " , bytes(msg.CorrelationId))'''
                    print ("Label: %s" % msg.Label)
                    print ("Recieve: %s \n" % msg.Body)
                    logger.info('Receive message: %s ' % msg.Label)

            return  msg.Id, msg.Label, msg.Body 

        myq.Close()

    #except Exception as e:
        #tkMessageBox.showerror("Error","Receive Error: %s" % str(e))
        #logger.error('Receive message, %s' % str(e))

def send(mId,lmsgs,msgs,pathname):
    #try:
        ourQueue = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
        ourQueue.FormatName = 'DIRECT=OS:'+pathname 
        respQueue = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
        respQueue.FormatName =  'DIRECT=OS:'+ glbconfig_(hname,'replyoi-msqueue') 
        resptQueue = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
        resptQueue.FormatName =  'DIRECT=OS:'+ glbconfig_(hname,'cc-msqueue') 

        try: 
            resq = respQueue.Open(2,0)
            
        except:
            msmq = gencache.EnsureModule('{D7D6E071-DCCD-11D0-AA4B-0060970DEBAE}', 0, 1, 0)
            resq = msmq.MSMQQueueInfo()
            resq.PathName = glbconfig_(hname,'replyoi-msqueue')
            resq.Create() 
            logger.info('Create message queue: %s' % resq.PathName) 

        mymq = ourQueue.Open(2,0)
        msg = win32com.client.Dispatch("MSMQ.MSMQMessage")
        msg.Body = str(msgs)
        msg.Label = str(lmsgs)

        if str(msg.Label) == 'SetupEnded':
            #resptQueue = win32com.client.Dispatch("MSMQ.MSMQDestination")
            #resptQueue.FormatName =  'DIRECT=OS:'+ glbconfig_(hname,'cc-msqueue')
            #msg.ResponseDestination = resptQueue
            print("Send Correlation ID: ", str(mId),type(mId))
            msg.CorrelationId = mId
        if pathname == glbconfig_(hname,'oi-msqueue'):
            respQueue = win32com.client.Dispatch("MSMQ.MSMQDestination")
            respQueue.FormatName =  'DIRECT=OS:'+ glbconfig_(hname,'replyoi-msqueue')
            msg.ResponseDestination = respQueue
        
        msg.Send(mymq)
        print ("Label : %s" % msg.Label)
        print ("Send : %s \n" % msg.Body)
        logger.info('Send message: %s  |  %s \n' % (msg.Label,msg.Body))
        mymq.Close()
        

    #except Exception as e:
        #tkMessageBox.showerror("Error","Send Error: %s" % str(e))
        #logger.error('Send message, %s' % str(e))

def objmsg(msgs,attrib):
    try:
        root = ET.fromstring(msgs)
        for key in root.findall('Dictionary/V'):
            ele = key.text
            atb = key.get('key')
            if atb == attrib:
                return ele
    except Exception as e:
        tkMessageBox.showerror("Error","Error: %s" % str(e))
        logger.error('Get message parameter ,',str(e))

   
def setupid(d):
    try:
        if not 'SetupId' in str(d):
            print("no setupid \n")
            tkMessageBox.showerror("Error","SetupID Error: %s" % str(e))
            logger.error('Parsing no setupID %s' % str(e))

        setupID =[]
        for setupid in d['response']['pdc']['Step']['SetupOptions']['SetupOption']: 
            for HardwareSetup in setupid['HardwareSetups']['HardwareSetup']:
                chnlmap = str(HardwareSetup['ChannelMapType'])
                if HardwareSetup['SetupId'].upper() not in setupID:
                    setupID.append(HardwareSetup['SetupId'])
                    print('HWSetup : %s' % HardwareSetup['SetupId'])
                    #print(setupID)
            break
        print("Setup ID: %s \n" % setupID)
        
        if len(setupID) == 1:
            #logger.info('Parsing SetupID: %s, %s'% (len(setupID), setupID))
            return len(setupID),chnlmap
        else:
            #logger.info('Parsing SetupID: %s, %s'% (len(setupID), setupID))
            return len(setupID),chnlmap
            
            
    except Exception as e:
        tkMessageBox.showerror("Error","SetupID Error: %s" % str(e))
        logger.error('Parsing setupID %s' %str(e))

def http_get(url, params):

    try:
        print('URL : %s?j=%s \n' %(url,params))
        f = urllib2.urlopen('%s?j=%s' %(url,params))
        data = f.read()
        #print('http_get : %s' % data)
        f.close()
        
        
    except urllib2.URLError:
        print('error : \n')
        tkMessageBox.showinfo("Error","Oops, we encountered an error while accessing, " )
        logger.error('Encountered an errorwhile accessing %s' % url) 

    return data
    
def callws(msgs):
    try:
        url = glbconfig_('URL', 'udal-url')
        params = '{"api":"setup","class":"OI","params":{"lotid":"'+str(objmsg(msgs,'LotID'))+'","tester_type":""}}'
        #print(url,params)
        response = http_get(url, params)
        if response:
            print('Loading json... \n')
            d = json.loads(response)
            logger.info('UDAL ws response: %s \n' % str(d))

            #check mpc, environment, testmode
            mpc = str(objmsg(msgs,'ProductID'))
            udal_mpc = str(d['response']['pdc']['MPC'])

            currentstep = d['response']['pdc']['CurrentStep']
            ccstep = currentstep[0:3]

            if '@' in str(objmsg(msgs,'Environment')): 
                envir = []
                envir = str(objmsg(msgs,'Environment')).split("@")
                envtemp = envir[1]
                #ccstep = envir[0]
            else: 
                envtemp = str(objmsg(msgs,'Environment'))

            if not 'TestStepTemp' in str(d):
                print('curentstep : %s \n' % str(currentstep))
                #tkMessageBox.showerror("Error","The current step is : ") #% str(currentstep))
                reply = ('<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">' + 
                        '<Dictionary key="Top">' +
                        '<V dt:dt="i4" key="ReturnCode">-1</V>' + 
                        '<V dt:dt="string" key="ReturnText">Current step Error </V>' + 
                        '<V dt:dt="string" key="ReturnDetails">The current step is ' + currentstep + '</V>' +
                        '<V dt:dt="string" key="SenderID">' + hname + '</V>' + 
                        '</Dictionary>' + 
                        '</Root>')     
                logger.error('The current step is : %s' % str(currentstep))
                return 'Cell Controller',reply,ccstep

            udal_envtemp = str(d['response']['pdc']['Step']['TestStepTemp'])
            testmode = str(objmsg(msgs,'TestMode'))

            if len(mpc) == 14 or len(udal_mpc) == 14:
                submpc = mpc[0:12]
                subudal_mpd = udal_mpc[0:12]
            else: 
                submpc = mpc 
                subudal_mpd = udal_mpc

            if submpc != subudal_mpd:
                print('Compare mpc: %s , %s, %s \n' % (mpc,udal_mpc,hname))
                #messagebox.showerror('MP Code does not match. Cell Controller =' +str(objmsg(msgs,'ProductID')) + \
                                       #', UDAL = ' + str(d['response']['pdc']['MPC']))
                reply = ('<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">' + 
                        '<Dictionary key="Top">' + 
                        '<V dt:dt="i4" key="ReturnCode">-1</V>' + 
                        '<V dt:dt="string" key="ReturnText">MP Code does not match</V>' + 
                        '<V dt:dt="string" key="ReturnDetails">Cell controller MPC = '+ mpc + ' and UDAL MPC = '+ udal_mpc +'</V>' + 
                        '<V dt:dt="string" key="SenderID">' + hname + '</V>' + 
                        '</Dictionary>' + 
                        '</Root>')       
                logger.error('MP Code does not match. Cell Controller = %s, UDAL = %s' % (mpc,udal_mpc))
                return 'Cell Controller',str(reply),str(ccstep)

            elif envtemp != udal_envtemp:
                print('Compare Environment: %s , %s, %s' % (envtemp,udal_envtemp,hname))
                #messagebox.showerror('Environment does not match. Cell Controller =' +str(objmsg(msgs,'Environment')) + \
                                    #   ', UDAL = ' + str(d['response']['pdc']['Step']['TestStepTemp']))
                reply = ('<Root xmlns:dt="urn:schemas-microsoft-com:datatypes"> ' + 
                        '<Dictionary key="Top">' +  
                        '<V dt:dt="i4" key="ReturnCode">-1</V> ' + 
                        '<V dt:dt="string" key="ReturnText">Environment does not match</V> ' + 
                        '<V dt:dt="string" key="ReturnDetails">Cell controller = ' + envtemp + ' and UDAL TEMP = '+ udal_envtemp +'</V>' +  
                        '<V dt:dt="string" key="SenderID">'  + hname + '</V> ' + 
                        '</Dictionary>' +  
                        '</Root>' )    
                logger.error('Environment does not match. Cell Controller = %s, UDAL = %s' % (envtemp,udal_envtemp))
                return 'Cell Controller',str(reply),str(ccstep)
            
            elif testmode == "FT_RANDOMQC" or testmode == "FT_QC" or testmode == "QC" and \
                ('FS' in currentstep or 'RS' in currentstep):
                #messagebox.showerror(Wrong test mode selection)
                print('Wrong test mode selection : %s , %s \n' % (testmode,currentstep))
                reply = ('<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">' +  
                        '<Dictionary key="Top">' +  
                        '<V dt:dt="i4" key="ReturnCode">-1</V>' + 
                        '<V dt:dt="string" key="ReturnText">Wrong test mode selection</V>' +  
                        '<V dt:dt="string" key="ReturnDetails">Cell controller = ' + testmode + ' and Currenstep ' + currentstep +'</V>' + 
                        '<V dt:dt="string" key="SenderID">' + hname + '</V>' +  
                        '</Dictionary>' +  
                        '</Root>')       
                logger.error('Wrong test mode selection. Cell Controller = %s, UDAL = %s' % (testmode,currentstep))
                return 'Cell Controller',str(reply),str(ccstep) 

            else:
                print('MPC and Environment is match \n')
                num,channelmaptype = setupid(d)
                print('num : %s, chlmap : %s \n' % (num,channelmaptype))

                if num > 1:
                    print('Not revise message, currentstep : %s \n' % str(ccstep))
                    return 'Cell Controller',str(msgs),str(ccstep)

                elif num == 1:
                    udal_channelmaptype = channelmaptype
                    HardwareMap =  str(objmsg(msgs,'HardwareMap'))
                    TestProgName =  str(objmsg(msgs,'TestProgFileName'))
                    TestProgChksum = str(objmsg(msgs,'TestProgChecksum'))
                    TestFlow = str(objmsg(msgs,'TestFlow'))
                    PartNum = str(objmsg(msgs,'PartNum'))
                    DeviceType = str(objmsg(msgs,'DeviceType'))
                    for SetupOption in d['response']['pdc']['Step']['SetupOptions']['SetupOption']: 
                        print(str(SetupOption['TesterType']))
                        for HardwareSetup in SetupOption['HardwareSetups']['HardwareSetup']:
                            if udal_channelmaptype == str(HardwareSetup['ChannelMapType']):
                            #if str(SetupOption['TesterType']) == glbconfig_('SetUpTester', 'tester-type'):
                                #revise testProgramName
                                print(str(SetupOption['TestProgMainSource']))
                                udal_testprogmains = str(SetupOption['TestProgMainSource']) 
                                #revise testProgChksum
                                print(str(SetupOption['TestProgChecksum']))
                                udal_testprogchksum = str(SetupOption['TestProgChecksum'] )
                                #revise testProgExcutable
                                print(str(SetupOption['TestProgExecutable']))
                                udal_testprogexcut = str(SetupOption['TestProgExecutable'] ) 
                                #revise Partnum Device
                                udal_Device = str(SetupOption['Device'])

                        break
                    #revise    
                    revise = msgs.replace(HardwareMap,udal_channelmaptype)
                    revise = revise.replace(TestProgName,'P:\\%s\\%s' % (udal_testprogmains[0:5],udal_testprogmains))
                    revise = revise.replace(TestProgChksum,udal_testprogchksum)
                    revise = revise.replace(TestFlow,udal_testprogexcut)
                    revise = revise.replace(PartNum,udal_Device)

                    if  str(d['response']['pdc']['Step']['ProgramType']) == None:
                        udal_ProgramType = ''
                        revise = revise.replace(DeviceType,udal_ProgramType)
                    else:
                        udal_ProgramType = str(d['response']['pdc']['Step']['ProgramType'])
                        revise = revise.replace(DeviceType,udal_ProgramType)    
                    #print('Revise: '+ revise) 
                    
                    logger.info('Revised message: %s' % revise)


                    return 'Web Service',str(revise),str(ccstep)

    except Exception as e:
        d = {}
        tkMessageBox.showerror("Error","Problem loading/parsing JSON data: %s" % str(e))
        logger.error('Problem loading/parsing JSON data, %s' % str(e))

      
def callUpdateProgWS(datasource,messge,ccurrentstep):
    try:
        url =glbconfig_('URL','updateTestProg-url')

        # structured 
        param = "PRODUCT_ID="+ str(objmsg(messge,'ProductID')) + "&"  \
                "LOT_ID=" + str(objmsg(messge,'LotID')) + "&"  \
                "TESTER_ID="+ hname  + "&"  \
                "HANDLER_ID=" + str(objmsg(messge,'HandlerID')) + "&" \
                "TEST_PROGRAM=" + str(objmsg(messge,'TestProgFileName')) + "&" \
                "TEST_PROGRAM_CHECKSUM=" + str(objmsg(messge,'TestProgChecksum'))+ "&" \
                "JOB_NAME=" + str(objmsg(messge,'TestFlow')) + "&" \
                "PART_NUMBER=" + str(objmsg(messge,'PartNum')) + "&" \
                "CHANNELMAP=" + str(objmsg(messge,'HardwareMap')) + "&" \
                "ENVIRONMENT=" + str(objmsg(messge,'Environment')) + "&" \
                "PROGRAMMING_TYPE=" + str(objmsg(messge,'DeviceType')) + "&" \
                "CP_ON=" + str(objmsg(messge,'CPOnChecksum')) + "&" \
                "CP_OFF=" + str(objmsg(messge,'CPOffChecksum')) + "&" \
                "Q_CODE=" + str(objmsg(messge,'QCode')) + "&" \
                "DATA_SOURCE=" + datasource + "&" \
                "TEST_MODE=" + str(objmsg(messge,'TestMode')) + "&" \
                "MES_TEST_STEP=" + ccurrentstep + "&" \
                "OI_VERSION=" + glbconfig_('SetupTester','software-version')

        print("URL: %s%s \n" %(url,param))
        resp = urllib2.request.urlopen('%s?%s' % (url,param))
        reply = resp.read()
        print("CallUpdateProgWS Reply: %s " % reply)
        resp.close()
       
        logger.info('Call web service UpdateTestProgramInfo response: %s' % reply)

        return reply

    except urllib2.URLError:
        tkMessageBox.showerror("Error","Oops, we encountered an error while accessing %s" % url)
        logger.error('Encountered an errorwhile accessing %s' % url)


tester = []
tstname = glbconfig_(hname,'oi-msqueue')
tester = tstname.split("\\")
print('Tester : %s \n' % str(tester[0]).upper())


def job2():
    try:
        l1.config(text = 'Autoload tool Version     %s' % glbconfig_("SetupTester","software-version"))
        l2.config(text = 'Tester                 %s' % str(tester[0]).upper() )
        l3.config(text = 'Status :                  Active   ')

        msgid,label,msgs = receive(glbconfig_(hname,'au-msqueue'))   #Receive from Cell controller
        if str(label) == 'SetupTester' and str(msgs) is not None :
            l3.config(text = 'Status : Start Lot')
            l3.config(text = 'Status : Recieved [%s]' % str(label))
            datasource,revisemsg,ccstep = callws(str(msgs))  #Call UDAL ws
            if 'Environment does not match' in str(revisemsg)  or  'MP Code does not match' in str(revisemsg) or  'Current step Error' in str(revisemsg) or 'Wrong test mode selection' in str(revisemsg):
                l3.config(text = 'Status : Revise failed')
                send(msgid,'SetupTesterReply',str(revisemsg),glbconfig_(hname,'cc-msqueue'))
                l3.config(text = 'Status : Sending [%s]' % 'SetupTesterReply')
            else:
                send(msgid,'SetupTester',str(revisemsg),glbconfig_(hname,'oi-msqueue'))
                l3.config(text = 'Status : Sending [%s]' % str(label))

        elif str(label) != '' and str(msgs) is not None:
            l3.config(text = 'Status : Recieved [%s]' % str(label)) 
            send(msgid,str(label),str(msgs),glbconfig_(hname,'oi-msqueue'))
            l3.config(text = 'Status : Sending [%s]' % str(label))
            if 'SetupEndedReply' in str(label) and str(msgs) is not None :   # Call Update Test Program ws
                updatetestprog = callUpdateProgWS(str(datasource),str(revisemsg),str(ccstep))
                print("updatetestprog reply: %s \n" % updatetestprog) 
    
        msgidoi,labeloi,msgsoi = receive(glbconfig_(hname,'replyoi-msqueue')) #Receive from Tester
        if str(labeloi) != '' and msgsoi is not None :
            l3.config(text = 'Status :  Recieved [%s]' % str(labeloi))
            send(msgid,str(labeloi),str(msgsoi),glbconfig_(hname,'cc-msqueue'))
            l3.config(text = 'Status :  Sending [%s]' % str(labeloi))
        
        threading.Timer(2.0,job2).start() 
        
    except Exception as e:
        logger.error('%s' % str(e))
        #tkMessageBox.showerror("Error",str(e))



    
app = tk.Tk()
app.title('Autoload tool')
app.geometry("440x130")
app.eval('tk::PlaceWindow . center')

l1 = tk.Label(app,font =("Courier", 12))
l2 = tk.Label(app,font =("Courier", 12))
l3 = tk.Label(app,font =("Courier", 12))

l1.pack(padx= (30,0),pady=(15,0),anchor ='w')
l2.pack(padx= (30,0),pady=(5,0),anchor ='w')
l3.pack(padx= (30,0),pady=(5,0),anchor ='w')
job2()

app.mainloop()

