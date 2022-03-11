import win32com.client
import os
from datetime import datetime, timedelta
outlook =win32com.client.Dispatch('outlook.application')
mapi = outlook.Getnamespace("Mapi")
for account in mapi.Account:
    messages = inbox.Items
received_dt = datetime.now() - timedelta(days=1)
received_dt = received_dt.strfttime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
messages = messages.Restrict("[SenderEmailAddress] = 'xrlab@derby.ac.uk'")
print(msg.Subject)