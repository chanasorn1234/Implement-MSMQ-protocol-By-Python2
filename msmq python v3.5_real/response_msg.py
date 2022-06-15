from time import sleep
import win32com.client
import os
import struct
from time import sleep
import pythoncom
import platform
import ConfigParser
import urllib2 
import sys
import json
import xml.etree.ElementTree as ET

type_setuptester = None
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
hname = hostname()

def receive(typesetup,msgs,pathname):#mId,typesetup,msgs,pathname):
    qinfo_id=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    keepid = win32com.client.Dispatch("MSMQ.MSMQMessage")
    pre_id = win32com.client.Dispatch("MSMQ.MSMQQueue")

    
    computer_name = os.getenv('COMPUTERNAME')
    pathname = computer_name+"\\PRIVATE$\\myqueue"
    qinfo_id.FormatName="DIRECT=OS:"+pathname
    pre_id = qinfo_id.Open(1,0)
    keepid = pre_id.Receive()
    # print(keepid)

    ############
    if(keepid.Label == 'SetupTester'):
        if(typesetup == 1):
            res_destination = win32com.client.Dispatch("MSMQ.MSMQDestination")
            resmsg = win32com.client.Dispatch("MSMQ.MSMQMessage")

            res_destination = keepid.ResponseDestination
            resmsg.Body = msgs#'<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">'+\
                # '<Dictionary key="Top">'+\
                # '<V dt:dt="i4" key="ReturnCode">0</V>'+\
                # '<V dt:dt="string" key="ReturnText"></V>'+\
                # '<V dt:dt="string" key="ReturnDetails"></V>'+\
                # '<V dt:dt="string" key="SenderID">'+hostname()+'</V>'+\
                # '</Dictionary>'+\
                # '</Root>'
            resmsg.Label = 'SetupTesterReply'

            resmsg.CorrelationId = keepid.Id
            
            print(type(keepid.Id))
            print(struct.unpack("<HH",resmsg.CorrelationId[16:20]),struct.unpack("<HH",keepid.Id[16:20]))
            
            resmsg.Send(res_destination)
            res_destination.Close()
            pre_id.Close()
        elif(typesetup == 2):
            res_destination = win32com.client.Dispatch("MSMQ.MSMQDestination")
            resmsg = win32com.client.Dispatch("MSMQ.MSMQMessage")

            res_destination = keepid.ResponseDestination
            resmsg.Body = msgs#'<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">'+\
                # '<Dictionary key="Top">'+\
                # '<V dt:dt="i4" key="ReturnCode">-1</V>'+\
                # '<V dt:dt="string" key="ReturnText">Program Not Loaded</V>'+\
                # '<V dt:dt="string" key="ReturnDetails"></V>'+\
                # '<V dt:dt="string" key="SetupTesterMsgID">{CDC5E747-EE4F-4A9A-A8E4-7F7F8709FAC4}\\23031291</V>'+\
                # '<V dt:dt="string" key="SenderID">'+hostname()+'</V>'+\
                # '</Dictionary>'+\
                # '</Root>'
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



def objmsg(msgs,attrib):
    try:
        root = ET.fromstring(msgs)
        for key in root.findall('Dictionary/V'):
            ele = key.text
            atb = key.get('key')
            if atb == attrib:
                return ele
    except Exception as e:
        print('error')
        # tkMessageBox.showerror("Error","Error: %s" % str(e))
        # logger.error('Get message parameter ,',str(e))
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
def setupid(d):
    try:
        if not 'SetupId' in str(d):
            print("no setupid \n")
            # tkMessageBox.showerror("Error","SetupID Error: %s" % str(e))
            # logger.error('Parsing no setupID %s' % str(e))

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
        print('Error SetupID')
        # tkMessageBox.showerror("Error","SetupID Error: %s" % str(e))
        # logger.error('Parsing setupID %s' %str(e))

def http_get(url, params):

    try:
        print('URL : %s?j=%s \n' %(url,params))
        f = urllib2.urlopen('%s?j=%s' %(url,params))
        data = f.read()
        #print('http_get : %s' % data)
        f.close()
        
        
    except urllib2.URLError:
        print('error : \n')
        # tkMessageBox.showinfo("Error","Oops, we encountered an error while accessing, " )
        # logger.error('Encountered an errorwhile accessing %s' % url) 

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
            # logger.info('UDAL ws response: %s \n' % str(d))

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
                # logger.error('The current step is : %s' % str(currentstep))
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
                # logger.error('MP Code does not match. Cell Controller = %s, UDAL = %s' % (mpc,udal_mpc))
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
                # logger.error('Environment does not match. Cell Controller = %s, UDAL = %s' % (envtemp,udal_envtemp))
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
                # logger.error('Wrong test mode selection. Cell Controller = %s, UDAL = %s' % (testmode,currentstep))
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
                    
                    # logger.info('Revised message: %s' % revise)


                    return 'Web Service',str(revise),str(ccstep)

    except Exception as e:
        d = {}
        print('Error Problem loading/parsing JSON data')
        # tkMessageBox.showerror("Error","Problem loading/parsing JSON data: %s" % str(e))
        # logger.error('Problem loading/parsing JSON data, %s' % str(e))
typesetup = 1
while(1):
    pre_qinfo_id=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    pre_keepid = win32com.client.Dispatch("MSMQ.MSMQMessage")
    pre_pre_id = win32com.client.Dispatch("MSMQ.MSMQQueue")

    precheck_computer_name = os.getenv('COMPUTERNAME')
    pre_pathname = precheck_computer_name+"\\PRIVATE$\\myqueue"
    pre_qinfo_id.FormatName="DIRECT=OS:"+pre_pathname
    pre_pre_id = pre_qinfo_id.Open(1,0)

    timeout_sec = 1.0
    check = pre_pre_id.peek(pythoncom.Empty, pythoncom.Empty, timeout_sec * 1000)
    if(check != None):
        # msgid,label,msgs = receive(glbconfig_(hname,'au-msqueue'))
        if(check.Label == 'SetupTester' and str(check.Body) is not None):
            datasource,revisemsg,ccstep = callws(str(check.Body))
            if 'Environment does not match' in str(revisemsg)  or  'MP Code does not match' in str(revisemsg) or  'Current step Error' in str(revisemsg) or 'Wrong test mode selection' in str(revisemsg):
                type_setuptester = 1
                receive(type_setuptester,str(revisemsg),'') ######ล่าสุดถึงตรงนี้อย่าลืมลบด้วย
            else:
                type_setuptester = 2
                receive(type_setuptester,str(revisemsg),'')
        else:
            sleep(5)
            pre_pre_id.Receive()

    sleep(5)
    type_setuptester = None

    
