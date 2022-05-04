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
Junk = mapi.GetDefaultFolder(23)
xrlabsFolder = mapi.GetDefaultFOlder(6).Folders["xrlabs"]
while True:
    for message in list(messages)[:1]:
        if message.SenderEmailAddress == 'xrlab@derby.ac.uk':
            dates = re.findall(r'(\d+/\d+/\d+)', message.Body)
            times = re.findall(r"(?i)(\d?\d:\d\d)", message.Body)
            dateofbooking = dates[0]
            enddate = dates[1]
            starttime = times[2]
            endtime = times[3]
            FMT = '%H:%M'
            tdelta = datetime.strptime(endtime, FMT) - datetime.strptime(starttime, FMT)
            time = str(tdelta)
            h,m,s = time.split(':')
            duration = int(h) * 60 + int(m)
            li = message.Body
            linesplit = li.splitlines()
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
            appt.Subject = Name + ' ' + Station + ' ' + booking_reference
            appt.Location = Room
            appt.Duration = duration
            appt.Save()
            message.Move(xrlabsFolder)