from bustime import StopMonitor
import string
from pymsgbox import *
import time

import win32com.client
from win32com.client import Dispatch, constants


starttime=time.time()

BUSTIME_API_KEY = ('e5b2ff54-4072-4d8d-bb85-9106bd1fb7b4')
monitor = StopMonitor(BUSTIME_API_KEY, '404231', 'x27', 1)

temp123 = str(monitor)
tempx = temp123.splitlines()

stopname = tempx[0]

tmp1 = tempx[1].split(' ')

numstops = tmp1[1]

tmp2 = tmp1[2].split('/')
tmp3 = tmp2[1].split('mi')

distance = tmp3[0]
busname = tmp1[0]

#print (busname)
#print (stopname)
#print (numstops)
#print (distance)

#print (time.time())

if float(distance) < 0.20:
    if float(numstops) < 1:
        print('COME OUTSIDE NOW')
        print('COME OUTSIDE NOW')
        print('COME OUTSIDE NOW')
        const=win32com.client.constants
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "COME OUTSIDE!!!"
        newMail.Body = "BUS IS HERE!!! X27"
        newMail.To = "718000000@mms.att.net;vz000000.com"
        newMail.send
        alert(text='COME OUTSIDE NOW', title='BUS ALERT', button='OK')
        alert(text='COME OUTSIDE NOW', title='BUS ALERT', button='OK')
        alert(text='COME OUTSIDE NOW', title='BUS ALERT', button='OK')
 
while float(distance) > 0.20:
    time.sleep(10.0 - ((time.time() - starttime) % 10.0))
    monitor = StopMonitor(BUSTIME_API_KEY, '404231', 'x27', 1)
    temp123 = str(monitor)
    tempx = temp123.splitlines()
    stopname = tempx[0]
    tmp1 = tempx[1].split(' ')
    numstops = tmp1[1]
    tmp2 = tmp1[2].split('/')
    tmp3 = tmp2[1].split('mi')
    distance = tmp3[0]
    busname = tmp1[0]
    if float(distance) < 0.20:
        if float(numstops) < 1:
            print('COME OUTSIDE NOW')
            print('COME OUTSIDE NOW')
            print('COME OUTSIDE NOW')
            const=win32com.client.constants
            olMailItem = 0x0
            obj = win32com.client.Dispatch("Outlook.Application")
            newMail = obj.CreateItem(olMailItem)
            newMail.Subject = "COME OUTSIDE!!!"
            newMail.Body = "BUS IS HERE!!! X27"
            newMail.To = "718000000@mms.att.net;vz000000.com"
            newMail.send
            alert(text='COME OUTSIDE NOW', title='BUS ALERT', button='OK')
            alert(text='COME OUTSIDE NOW', title='BUS ALERT', button='OK')
            alert(text='COME OUTSIDE NOW', title='BUS ALERT', button='OK')
