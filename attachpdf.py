import win32com.client as win32
import os, fnmatch



BASE_DIR = os.path.dirname(os.path.abspath(__file__))
files = os.listdir(BASE_DIR)

pdffiles = [file for file in files if fnmatch.fnmatch(file,'*.pdf')]
attachfiles = list(map(lambda x: os.path.join(BASE_DIR,x),pdffiles))


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
#mail.To = RepEmail  #change this line to change receipient's emails
#mail.To = 'accounting@stingerchemicals.com'  #change this line to change receipient's emails
#mail.CC = CCEmails
#mail.Subject = "Test attach pdf"

mail.HTMLBody = "<h1>HELLO WORLD</h1>"
print ('There are {0} files'.format(len(attachfiles)))

for each in attachfiles:
	mail.Attachments.Add(each)
	print (f'{each} is attached')
mail.Display()
