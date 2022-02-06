import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

myFolder = outlook.GetDefaultFolder(6)

mySubFolder = myFolder.Folders(1)

mySubSubFolder = mySubFolder.Folders(1)

print (myFolder)

print (mySubFolder)

print (mySubSubFolder)

messages = mySubSubFolder.Items
message = messages.GetLast()

#body_content = message.body
#print (body_content)

i = 0
while i < 10:
    print(message.body.encode("utf-8"))
    message = messages.GetPrevious()
    i = i + 1
