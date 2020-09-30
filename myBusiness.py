import myExcel
import myWord
import myQRCode
import myMail
import myConfig
import myGUI
#qrimagefile = excel.qrcode_create('http://34.82.12.150/')

#wordfile = excel.word_crate_from_template_with_qrcode('./~tempfile.docx',qrimagefile)

#excel.email_send("lucasain2010@gmail.com",'./~tempfile.docx')


# from multiprocessing import Process
import threading, sys, os
import time
# import signal

if_true_exit = False

def RepresentsInt(s):
    try: 
        int(s)
        return True
    except ValueError:
        return False
time_to_start = 1
time_after_run = 20
def get_delay():
    print("""
    
    GET DELAY
    
    """)
    if myConfig.configuration_file_has_been_persisted():
        print("has been persisted")
        
        #print(myConfig.check_data("EMAIL",'subject'))

        # for key in (myConfig.get_mail_data()): print(key)
        
        if myConfig.check_data('Automate outlook mailing',"time_to_run"):
            print("time_to_run exist")

            data_raw = myConfig.get_saved_data('Automate outlook mailing',"time_to_run")
            if RepresentsInt(data_raw):
                print("time_to_run Okey")
                time_to_start = int(data_raw)
                print("""
                #####################################
                """)
                print(time_to_start)
                print("""
                #####################################
                """)
                pass
            else:
                print("time_to_run ValueError")
                #print(time_to_start)
                pass
        else:
            print("time_to_run not exist")
            myConfig.configuration_file_create_persist()
            pass
        
        #myExcel.getDataFromExcel(excel_file,)
        pass
    else:
        print("The program needs to be configured")
        myConfig.configuration_file_create_persist()
        pass
    pass

def get_after_delay():
    print("""
    
    GET DELAY
    
    """)
    if myConfig.configuration_file_has_been_persisted():
        print("has been persisted")
        
        #print(myConfig.check_data("EMAIL",'subject'))

        # for key in (myConfig.get_mail_data()): print(key)
        
        if myConfig.check_data('Automate outlook mailing',"time_after_run"):
            print("time_after_run exist")

            data_raw = myConfig.get_saved_data('Automate outlook mailing',"time_after_run")
            if RepresentsInt(data_raw):
                print("time_after_run Okey")
                time_after_run = int(data_raw)
                print("""
                #####################################
                """)
                print(time_after_run)
                print("""
                #####################################
                """)
                pass
            else:
                print("time_after_run ValueError")
                #print(time_to_start)
                pass
        else:
            print("time_after_run not exist")
            myConfig.configuration_file_create_persist()
            pass
        
        #myExcel.getDataFromExcel(excel_file,)
        pass
    else:
        print("The program needs to be configured")
        myConfig.configuration_file_create_persist()
        pass
    pass

"""
qrimagefile = myQRCode.qrcode_create('http://34.82.12.150/')

wordfile = myWord.word_create_from_template('./~tempfile.docx',qrimagefile,"Luca","hello there","ABC123","09/09/1999","09/10/1999","Bye","luca@luca.com","luca company","0800-XXXX")

myMail.email_send("lucasain2010@gmail.com","test subject","test body body","<p>my test html body</p>",wordfile)
"""
"""
dataJson =  myExcel.excel_read('./~tempfile.xlsx')

def checkIfIsNoneType(some_expected_word):
    if some_expected_word is None:
        #thing_to_return = "none"
        thing_to_return = str(some_expected_word)
        pass
    else:
        thing_to_return = some_expected_word
        pass

    # if some_expected_word is not None:

    return thing_to_return

for thing in dataJson:
    print(thing)
    print(thing['First Name'])
    

    qrimagefile = myQRCode.qrcode_create(thing['Code'])

    wordfile = myWord.word_create_from_template('./~tempfile.docx',qrimagefile,str(thing["First Name"]) +" "+ str(thing["Last Name"]),thing['Greeting'],thing['Unit'],thing['Arrival'],thing['Departure'],thing['Salutation'],thing['Maddress'],thing['Company Name'],thing['Phone'])

    myMail.email_send("lucasain2010@gmail.com","test subject","test body body","<p>my test html body</p>",wordfile)

    time.sleep(30)
    pass
"""

excel_file = "~tempfile.1.xlsx"
#myExcel.getDataFromExcel(excel_file)
"""
if myConfig.configuration_file_has_been_persisted():
    print("The program starts")
    
    print(myConfig.check_data("EMAIL",'subject'))
    # for key in (myConfig.get_mail_data()): print(key)

    if myConfig.check_data("EMAIL",'subject'):
        print(myConfig.get_mail_data()['subject'])
        myConfig.get_mail_data()['subject']
        pass
    else:
        myConfig.configuration_file_create_persist()
        pass
    #myExcel.getDataFromExcel(excel_file,)
    pass
else:
    print("The program needs to be configured")
    myConfig.configuration_file_create_persist()
    pass
"""

# myConfig.configuration_file_create_persist()
"""
[Automate outlook mailing]
auto_run = True
time_to_run = "00:15"
scheduled_time = "21:00"
"""
def the_aim_of_the_program():
    if myConfig.configuration_file_has_been_persisted():
        print("The program starts")
        
        #print(myConfig.check_data("EMAIL",'subject'))

        # for key in (myConfig.get_mail_data()): print(key)
        
        if myConfig.check_data("EMAIL",'subject'):
            #print(myConfig.get_mail_data()['subject'])
            #myConfig.get_mail_data()['subject']

            print()
            print()
            print()

            theFile = myConfig.get_saved_data("EXCEL",'file')
            theSubject = myConfig.get_saved_data("EMAIL",'subject')
            theBody = myConfig.get_saved_data("EMAIL",'body')
            theBodyInHTML = "<p>" + myConfig.get_saved_data("EMAIL",'body') + "</p>"
            print("the aim")
            #print(theFile,theSubject,theBody,theBodyInHTML)
            myExcel.getDataFromExcel(theFile,theSubject,theBody,theBodyInHTML)
            pass
        else:
            print("subject not exist")
            myConfig.configuration_file_create_persist()
            pass
        #myExcel.getDataFromExcel(excel_file,)
        pass
    else:
        print("The program needs to be configured")
        myConfig.configuration_file_create_persist()
        pass
    pass

def my_func():
    print("delay starts")
    get_delay()
    time.sleep(time_to_start)

    print("delay passed")
    print("""
    #########################################
    Automated Process started
    #########################################
    """)
    the_aim_of_the_program()
    print("Process finished")
    """
    import sys
    sys.exit(1)
    exit()
    quit()
    """
    # get_after_delay()
    # time.sleep(time_after_run)

    myShutdown()

def myShutdown():
    """
    docstring
    """
    if_true_exit = True
    print("Process Shutdown")
    #import sys
    #sys.exit(1)
    #exit()
    #quit()

    # Works
    # import os
    os._exit(1)
    """
    import psutil

    current_system_pid = os.getpid()

    ThisSystem = psutil.Process(current_system_pid)
    ThisSystem.terminate()
    """
    # https://stackoverflow.com/questions/1489669/how-to-exit-the-entire-application-from-a-python-thread
    # threading.interrupt_main()
    pass


def the_aim_of_the_program_with_delay():
    if myConfig.configuration_file_has_been_persisted():
        print("The program starts with delay")
        
        #print(myConfig.check_data("EMAIL",'subject'))

        # for key in (myConfig.get_mail_data()): print(key)
        
        if myConfig.check_data('Automate outlook mailing',"auto_run"):
            print("delay exist")
            if myConfig.get_saved_data('Automate outlook mailing',"auto_run") == 'True':
                print("auto_run is True")
                # time.sleep(10)

                ############################
                print("Process starts")
                print("delay starts")
                get_delay()
                time.sleep(time_to_start)

                print("delay passed")
                print("""
                #########################################
                Automated Process started
                #########################################
                """)
                the_aim_of_the_program()
                print("""
                #########################################
                Automated Process finished
                #########################################
                """)
                get_after_delay()
                print("after delay starts")
                time.sleep(time_after_run)
                print("after delay passed")
                print("Process finished")
                

                ############################

                print("bye guy")
                myShutdown()
                
                pass
            else:
                print("auto_run is False")
                # myConfig.configuration_file_create_persist()
                pass
            pass
        else:
            print("auto_run not exist")
            myConfig.configuration_file_create_persist()
            pass
        
        #myExcel.getDataFromExcel(excel_file,)
        pass
    else:
        print("The program needs to be configured")
        myConfig.configuration_file_create_persist()
        pass
    pass
