import win32com.client
import re
import os
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNameSpace("Mapi")
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items
messages = messages.Restrict("[SenderEmailAddress] = 'xrlab@derby.ac.uk'")
messages.Sort("[ReceivedTime]", Descending = True)
for message in list(messages)[:1]:
    if message.SenderEmailAddress == 'xrlab@derby.ac.uk':
        dates = re.findall(r'(\d+/\d+/\d+)', message.Body)
        times = re.findall(r"(?i)(\d?\d:\d\d)", message.Body)
        dateofbooking = dates[0]
        starttime = times[2]
        endtime = times[3]
        li = message.Body
        
        linesplit = li.splitlines()
        print(linesplit)
        Booking = linesplit[13]
        Station = (Booking[44:70])
        Room = (Booking[78:90])
        BookingName = linesplit[12]
        N = BookingName.split("Hello")
        Name = N[1]
        bookingreference = linesplit[15]
        Reference = bookingreference.split("is")
        booking_reference = Reference[1]
        appt = outlook.CreateItem(1)
        appt.Start  = dateofbooking + ' ' + starttime
        appt.Subject = Station + ' ' + booking_reference
        appt.Location =  Name + ' ' + Room
        appt.Save()
        
        