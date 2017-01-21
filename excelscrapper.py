from openpyxl import load_workbook
from mailer.mailer_v3 import MailerApp
from sms.SMSMessage import Messenger
from docwriter.utkng_editor import EditIt


def run_defaulter_system(console, file):
    defaulter = []
    wb2 = load_workbook(file, read_only=True, data_only=True)
    sheet_name = wb2['defaulter']
    contact_sheet = wb2['contact']
    email_sheet = wb2['email']
    sms_message = "You are in defaulter list due to your attendance being less"

    mylog = open('log.txt', 'w')

    for x in range(11, 14):
        if sheet_name['R' + str(x)].value is not None:
            if sheet_name['R' + str(x)].value != "%":
                if int(sheet_name['R' + str(x)].value) < 75:
                    student_att_percent = str(sheet_name['R' + str(x)].value)
                    student_name = str(sheet_name['B' + str(x)].value)
                    student_roll = str(sheet_name['A' + str(x)].value)
                    student_email = email_sheet['D' + str(int(student_roll) + 1)].value  # update sheet
                    student_mobile = str(contact_sheet['D' + str(int(student_roll) + 1)].value)  # update sheet

                    editit = EditIt('docwriter/undertaking.docx', student_name, student_roll, student_att_percent)
                    fno = str(editit.read_n_write())

                    mail_client = MailerApp(student_email, "appuarunnair@gmail.com", "f1011995mon", "undertaking",
                                            student_att_percent, student_name, student_roll)

                    mail_client.htmladd("Attendance Report For {name} Roll No {roll} TECM".format(name=student_name,
                                                                                                  roll=student_roll))
                    mail_client.addattach(["temp.docx"])
                    mail_client.send()

                    messenger = Messenger(student_mobile, sms_message)
                    messenger.compose()

                    console.append("Defaulter : {name} {perc}% \n".format(name=student_name, perc=student_att_percent))
                    defaulter.append(student_name)
                    mylog.write("Name:{name}\nMobile:{mobile}\nEmail:{email}\n\n".format(name=student_name,
                                                                                         mobile=student_mobile,
                                                                                         email=student_email))

    mylog.close()
    '''f = open('defaulter_names.txt', 'w')
    for names in defaulter:
        f.write("" + "\n\n")
    f.close()'''