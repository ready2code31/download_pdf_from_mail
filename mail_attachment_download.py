import imaplib
import email
import os
import time
import win32api
import win32print
import PyPDF2
from dotenv import load_dotenv

i = 0
quote = '"'

load_dotenv(dotenv_path="config")

username = os.getenv("gmail_username")
password = os.getenv("gmail_password")

sv_dir = os.getenv("svdir")

ghostscript_path = os.getenv("ghostscript")
gsprint_path = os.getenv("gsprint")


def printfilelandscape(filename, ghostscript_path, gsprint_path):
    print("fonction printfile")
    
    currentprinter = win32print.GetDefaultPrinter()
    
    win32api.ShellExecute(0, 'open', gsprint_path, '-ghostscript "'+ ghostscript_path +'" -landscape"' +'" -printer "' + currentprinter + '" '+ quote +filename+ quote, '.', 0)


def printfileportrait(filename, ghostscript_path, gsprint_path):
    print("fonction printfile")
    
    currentprinter = win32print.GetDefaultPrinter()
    
    win32api.ShellExecute(0, 'open', gsprint_path, '-ghostscript "'+ ghostscript_path +'" -portrait"' +'" -printer "' + currentprinter + '" '+ quote +filename+ quote, '.', 0)


def orientationTest(filename):
    pdf = PyPDF2.PdfReader(open(filename, 'rb'))
    page = pdf.getPage(0).mediaBox

    if page.getUpperRight_x() - page.getUpperLeft_x() > page.getUpperRight_y() - page.getLowerRight_y():
        return 'Landscape'
    else:
        return 'Portrait'


def printbestorientationchoice(filename, ghostscript_path, gsprint_path):
    if orientationTest(filename) == 'Landscape':
        printfilelandscape(filename, ghostscript_path, gsprint_path)
    else:
        printfileportrait(filename, ghostscript_path, gsprint_path)



mail=imaplib.IMAP4_SSL("imap.gmail.com",993)
mail.login(username,password)
mail.select("Inbox")

typ, msgs = mail.search(None, '(SUBJECT "Bordereau")')
msgs = msgs[0].split()

#retrieve number of mails
i_boucle = len(msgs)

for emailid in msgs:

    #display in terminal the percentage of loading
    print(int((i/i_boucle)*100), "%")
    i=i+1
    resp, data = mail.fetch(emailid, "(RFC822)")
    raw_email = data[0][1] 

    raw_email_string = raw_email.decode('utf-8')
    m = email.message_from_string(raw_email_string)
    
    if m.get_content_maintype() != 'multipart':
        continue

    for part in m.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        
        filename=part.get_filename()
        if filename is not None:
            sv_path = os.path.join(sv_dir, filename)
            if not os.path.isfile(sv_path):     
                fp = open(sv_path, 'wb')
                fp.write(part.get_payload(decode=True))

                print(orientationTest(sv_path))
                filename_temp = filename
                time.sleep(3)
                print(filename_temp)
                #printbestorientationchoice(filename)
                    
                fp.close()

print("END")
