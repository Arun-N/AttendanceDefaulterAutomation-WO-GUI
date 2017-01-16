from docx import Document


class EditIt:

    file_no = 0

    def __init__(self, filename, stud_name, stud_rno, stud_percent):
        self.file_name = filename
        self.stud_name = stud_name
        self.stud_rno = stud_rno
        self.stud_percent = stud_percent
        self.count = 0

    def read_n_write(self):
        document = Document(self.file_name)
        for para in document.paragraphs:
            word = para.text.split()
            if not word:
                continue
            else:
                if word[0] == "Name:":
                    print("Got the name")
                    inline = para.runs
                    # Loop added to work with runs (strings with same style)
                    for i in range(len(inline)):
                        if 'Name:' in inline[i].text:
                            text = inline[i].text.replace('Name:', "Name: " + self.stud_name)
                            inline[i].text = text
                    self.count += 1
                elif word[0] == "Rollno:":
                    print("Got the RNO")
                    inline = para.runs
                    for i in range(len(inline)):
                        # check here
                        if 'Rollno:' in inline[i].text:
                            print("replaced")
                            text = inline[i].text.replace('Rollno:', "Rollno: " + self.stud_rno)
                            inline[i].text = text
                    self.count += 1
                elif word[0] == "Attendance:":
                    print("Got the Attd")
                    inline = para.runs
                    for i in range(len(inline)):
                        if 'Attendance:' in inline[i].text:
                            text = inline[i].text.replace('Attendance:', "Attendance: " + self.stud_percent)
                            inline[i].text = text
                    self.count += 1
                if self.count == 3:
                    EditIt.file_no += 1
                    document.save("temp" + str(EditIt.file_no) + ".docx")
                    #time.sleep(2)
                    self.count = 0
                    return EditIt.file_no

'''et = EditIt('undertaking.docx', "Arun Nair", "3232", "73")
et.read_n_write()'''