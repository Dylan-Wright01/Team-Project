from ctypes import addressof
import win32com.client
import re
import os
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('outlook.application') # This will connect to the outlook Application
mapi = outlook.GetNameSpace("Mapi")
inbox = mapi.GetDefaultFolder(6) # Connecting to the inbox of the email
messages = inbox.Items
messages.Sort("[ReceivedTime]", Descending = True) #Read the first emails in the Inbox
Junk = mapi.GetDefaultFolder(23)
xrlabsFolder = mapi.GetDefaultFolder(6).Folders["xrlabs"] # This is the subfolder in the inbox that the emails will be moved too
while True:
    for message in list(messages)[:10]:#This line of code reads the first 10 email in the inbox, this is so that the program will not have to read the entire inbox and instead read incoming emails
        try:
            if message.SenderEmailAddress == 'xrlab@derby.ac.uk':
                #Parsing the body of the email and saving the data
                dates = re.findall(r'(\d+/\d+/\d+)', message.Body) #Finds all the dates in the body of the email
                times = re.findall(r"(?i)(\d?\d:\d\d)", message.Body)#Finds all the times in the body of the email
                dateofbooking = dates[0]
                enddate = dates[1]
                starttime = times[2] 
                endtime = times[3]
                FMT = '%H:%M'
                tdelta = datetime.strptime(endtime, FMT) - datetime.strptime(starttime, FMT)
                time = str(tdelta)
                h,m,s = time.split(':')
                duration = int(h) * 60 + int(m)#THis line calculates the duration of the booking
                li = message.Body
                linesplit = li.splitlines()#This line splits each line in the body depending on the spaces 
                Booking = linesplit[13]
                Station = (Booking[44:70])#This line saves the characters betweek 44 and 70 as the station being used in the booking
                Room = (Booking[78:90])#This line saves the room being used
                BookingName = linesplit[12]
                N = BookingName.split("Hello")
                Name = N[1]#Saving the name of the booking
                bookingreference = linesplit[15]
                Reference = bookingreference.split("is")
                booking_reference = Reference[1]# Gets the booking reference of the booking
                appt = outlook.CreateItem(1)#Creating a calendar appointment
                appt.Start  = dateofbooking + ' ' + starttime
                appt.Subject = Name + ' ' + Station + ' ' + booking_reference
                appt.Location = Room
                appt.Duration = duration
                appt.Save()
                message.Move(xrlabsFolder)
                print("Email Successfuly Moved")
        
        except:
                print("An Error Occured")

        
        

    
        