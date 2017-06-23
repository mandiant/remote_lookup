#	Contributed by FireEye FLARE Team
#	Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
#   Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
#	License: GPL

import socket

class CRemote:
    
    def __init__(self, pc):
        self.ip = pc
        self.response = '' 
        self.debug = False
    
    def __del__(self):
        self.close()
        
    def __parse(self):
        if len(self.response) == 0 : return 0
        a = self.response.find(':')
        if a == -1 : return 0
        respCode = self.response[0:a]
        self.response = self.response[a+1:]
        if respCode == "ok" : return 1
        return 0;
    
    def close(self):
        try: 
            self.s.close()
        except:
            pass    
    
    def attach(self,processOrPID):
        self.response = ''
        
        try:
            processOrPID = str(processOrPID) # in case its a pid
            self.close()
            if self.debug: print("* Trying to connect to %s:9000" % self.ip)
            self.s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.s.settimeout(None)        
            self.s.connect((self.ip, 9000))
            if self.debug: print("* Trying to attach to " + processOrPID)
            self.s.send('attach:'+processOrPID+'\r') 
            self.response = self.s.recv(1024)
            if self.debug: print("* response received")        
        except Exception as e: 
            self.response = "fail:Exception %s" % e        

        return self.__parse()
    
    def resolve(self,lookup):
        self.response = ''
        
        try:
            if self.debug: print("* trying to resolve " + lookup)
            self.s.send('resolve:'+lookup+'\r')        
            self.response = self.s.recv(1024)
            if self.debug: print("* response received")
        except Exception as e: 
            self.response = "fail:Exception %s" % e        
            
        return self.__parse()    

remote = CRemote('192.168.0.67')
#remote.debug = 1

#example output:
#    Resolved ok: 7C80AE30 , GetProcAddress , 409 , kernel32.dll
#    attached to cmd 38 dlls 9773 export loaded time:3.359 seconds

if remote.attach('explorer'):
    if remote.resolve('getprocaddress'):
        print 'Resolved ok: ' + remote.response 
    else:
        print "Failed: " + remote.response 
else:
    print 'Failed to attach to process: ' + remote.response

if remote.attach(824): #attach by pid
    print 'attached to cmd ' + remote.response
else:
    print 'Failed to attach to pid'
