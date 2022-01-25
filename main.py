from Email.config import send_email

to_email = ""
from_email = ""
subject = ""
message = ""
cc = [""]
host = "smtp.gmail.com"  #or any other host of your choice
port = 587 #put the relevant port
username = ''
password = ""
files = ['./Email/files/test1.txt', './Email/files/test2.txt']

try:
    send_email(to_email, from_email, subject, message, cc, files, host, port, username, password)
except Exception as e:
    print(e)
