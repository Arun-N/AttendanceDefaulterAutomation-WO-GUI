from openpyxl import load_workbook
from mailer.mailer_v3 import MailerApp
from sms.SMSMessage import Messenger
from docwriter.utkng_editor import EditIt

# if __name__ == "main":

defaulter = []
wb = load_workbook('attendance.xlsx', read_only=True)
sheet_name = wb['attendance']
sms_message = "You are in defaulter list due to your attendance being less"

for x in range(2, 4):
    if float(sheet_name['E'+str(x)].value) < 75.0:
        print(float(sheet_name['E'+str(x)].value))
        student_att_percent = str(sheet_name['E'+str(x)].value)
        student_name = str(sheet_name['B'+str(x)].value)
        student_roll = str(sheet_name['A'+str(x)].value)
        student_email = sheet_name['C'+str(x)].value
        student_mobile = str(sheet_name['D'+str(x)].value)

        editit = EditIt('docwriter/undertaking.docx', student_name, student_roll, student_att_percent)
        fno = str(editit.read_n_write())

        # time.sleep(2)

        mail_client = MailerApp(student_email, "email_id", "password", "undertaking",
                                student_att_percent, student_name, student_roll)

        mail_client.htmladd("Attendance Report For {name} Roll No {roll} TECM".format(name=student_name,
                                                                                      roll=student_roll))
        mail_client.addattach(["temp" + fno + ".docx"])
        mail_client.send()

        messenger = Messenger(student_mobile, sms_message)
        messenger.compose()

        print("Defaulter : {name}".format(name=student_name))
        defaulter.append(student_name)

f = open('defaulter_names.txt', 'w')
for names in defaulter:
    f.write(names + "\n")
f.close()
