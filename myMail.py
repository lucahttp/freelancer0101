# import pythoncom
import win32com.client as win32
import os
import myConfig

# pythoncom.CoInitialize()
# pythoncom.CoInitialize()



#thread = threading.Thread(target=myBusiness.the_aim_of_the_program_with_delay, args=())
#thread.daemon = True                            # Daemonize thread
#thread.start()                                  # Start the execution
#my_func()



# import win32com.client as win32
from threading import Thread
import pythoncom

"""
class ExcelManip:
    def __init__(self):
    self.xlApp = win32com.client.Dispatch('Excel.Application')
    self.xlApp.Visible = True

def createExcel():
    import pythoncom
    pythoncom.CoInitialize()
    excel = ExcelManip()

thread = Thread(target = createExcel)
thread.start()
"""


"""
if __name__ == '__main__':
    start()
"""

def email_validation(x):
    a=0
    y=len(x)
    dot=x.find(".")
    at=x.find("@")
    for i in range (0,at):
        if((x[i]>='a' and x[i]<='z') or (x[i]>='A' and x[i]<='Z')):
            a=a+1
            pass
    if(a>0 and at>0 and (dot-at)>0 and (dot+1)<y):
        print("Valid Email")
        return True
    else:
        print("Invalid Email")
        return False


class EmailMaster:
    def __init__(self,recipient,subject,body=None,body_in_html=None,qrcode_attachement=None,file_attachement=None):
        if email_validation(recipient) == True:
            self.outlook = win32.Dispatch('outlook.Application')
            self.outlook.Visible = True
            #outlook = win32.Dispatch('outlook.application')
            #outlook.Visible = True
            mail = self.outlook.CreateItem(0)
            #mail.To = 'To address'
            mail.To = str(recipient)
            """
            mail.Subject = 'Message subject'
            mail.Body = 'Message body'
            mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
            """
            mail.Subject = str(subject)

            if body != None:
                mail.Body = str(body)
                pass
            if body_in_html != None:
                mail.HTMLBody = str(body_in_html)
                pass
            
            #this field is optional

            # To attach a file to the email (optional):
            #attachment  = "Path to the attachment"


            if qrcode_attachement != None:
                print(qrcode_attachement)
                # https://stackoverflow.com/questions/51520/how-to-get-an-absolute-file-path-in-python
                import os
                #mypath = os.path.abspath(str(qrcode_attachement))
                mypath = myConfig.getPathExternal(qrcode_attachement)
                print(mypath)
                #mail.Attachments.Add(mypath)

                attochment = mail.Attachments.Add(mypath)
                attochment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
                pass
            if file_attachement != None:            
                print(file_attachement)
                # https://stackoverflow.com/questions/51520/how-to-get-an-absolute-file-path-in-python
                import os
                #myfilepath = os.path.abspath(str(file_attachement))
                myfilepath = myConfig.getPathExternal(file_attachement)
                print(myfilepath)
                mail.Attachments.Add(myfilepath)
                pass

            mail.Send()
            # tested with outlook 2019 in windows 10 pro 64 bit
            pass
        else:
            "Email addres is unrecogn"
            pass



def email_send(recipient,subject,body=None,body_in_html=None,qrcode_attachement=None,file_attachement=None):
    if email_validation(recipient) == True:
        outlook = win32.Dispatch('outlook.application')
        # outlook.Visible = True
        mail = outlook.CreateItem(0)
        #mail.To = 'To address'
        mail.To = str(recipient)
        """
        mail.Subject = 'Message subject'
        mail.Body = 'Message body'
        mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
        """
        mail.Subject = str(subject)

        if body != None:
            mail.Body = str(body)
            pass
        if body_in_html != None:
            mail.HTMLBody = str(body_in_html)
            pass
        
        #this field is optional

        # To attach a file to the email (optional):
        #attachment  = "Path to the attachment"


        if qrcode_attachement != None:
            print(qrcode_attachement)
            # https://stackoverflow.com/questions/51520/how-to-get-an-absolute-file-path-in-python
            import os
            #mypath = os.path.abspath(str(qrcode_attachement))
            mypath = myConfig.getPathExternal(qrcode_attachement)
            print(mypath)
            #mail.Attachments.Add(mypath)

            attochment = mail.Attachments.Add(mypath)
            attochment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
            pass
        if file_attachement != None:            
            print(file_attachement)
            # https://stackoverflow.com/questions/51520/how-to-get-an-absolute-file-path-in-python
            import os
            #myfilepath = os.path.abspath(str(file_attachement))
            myfilepath = myConfig.getPathExternal(file_attachement)
            print(myfilepath)
            mail.Attachments.Add(myfilepath)
            pass

        mail.Send()
        # tested with outlook 2019 in windows 10 pro 64 bit
        pass
    else:
        "Email addres is unrecogn"
        pass
    pass


def OutlookMailer():
    # OutlookMailer.py
    # Python 2.7.6


    outlook = win32.Dispatch("Outlook.Application")

    """
    Source - https://msdn.microsoft.com/en-us/library/office/ff869291.aspx
    Outlook VBA Reference 
    0 - olMailItem
    1 - olAppointmentItem
    2 - olContactItem
    3 - olTaskItem
    4 - olJournalItem
    5 - olNoteItem 
    6 - olPostItem
    7 - olDistributionListItem
    """
    # mail = outlook.CreateItem(0X0)
    mail = outlook.CreateItem(0)

    mail.To = "mail1@example.com"

    mail.CC = "mail2@example.com"

    mail.BCC = "mail3@example.com"

    mail.Subject = "Test mail from Python"

    # Using "Body" constructs body as plain text
    # mail.Body = "Test mail body from Python"

    """
    Using "HtmlBody" constructs body as html text
    default font size for most browser is 12
    setting font size to "-1" might set it to 10
    """
    mail.HTMLBody = """
    <html>
    <head></head>
    <body>
        <font color="DarkBlue" size=-1 face="Arial">
        <p>Hi!<br>
        How are you?<br>
        Test HTML mail body from Python
        </p>
        </font>
    </body>
    </html>
    """

    """
    Set the format of mail
    1 - Plain Text
    2 - HTML
    3 - Rich Text
    """
    mail.BodyFormat = 2


    # Instead of sending the message, just display the compiled message
    # Useful for visual inspection of compiled message
    mail.Display(True)

    # Send the mail
    # Use this directly if there is no need for visual inspection
    # mail.Send()
    pass


def email_send_with_html_body(recipient,subject,body,body_in_html,attachement):
    if email_validation(recipient) == True:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        #mail.To = 'To address'
        mail.To = str(recipient)
        """
        mail.Subject = 'Message subject'
        mail.Body = 'Message body'
        mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
        """
        mail.Subject = str(subject)
        mail.Body = str(body)
        mail.HTMLBody = str(body_in_html)
        #this field is optional

        # To attach a file to the email (optional):
        #attachment  = "Path to the attachment"

        print(attachement)
        # https://stackoverflow.com/questions/51520/how-to-get-an-absolute-file-path-in-python
        mypath = os.path.abspath(str(attachement))
        print(mypath)
        mail.Attachments.Add(mypath)

        #mail.Attachments.Add(file_location + 'test.png')
        qrcode_html = "<img src='cid:"+attachement+"'> "

        mail.HTMLBody = "<html><body>  <img src='cid:"+attachement+"'>  </body></html>";
        
        mail.Send()
        # tested with outlook 2019 in windows 10 pro 64 bit
        pass
    else:
        "Email addres is unrecogn"
        pass
    pass
