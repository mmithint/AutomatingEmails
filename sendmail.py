import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate


import pandas as pd
def MailToProffesor(SNO, ProffesorName, ProffesorEmail, CourseNumber, CourseName):
    from_addr = 'mmithint@asu.edu'
    # "mohananitej98@gmail.com"
    to_addr = f"{ProffesorEmail}"
    subject = f'Application for Grader Position - {CourseNumber} : {CourseName}'
    content = f"""Dear Professor {ProffesorName},

I hope this email finds you well. I am writing to express my keen interest in the position of grader for the {CourseNumber} : {CourseName}. I am currently a Master's student in Computer Science at Asu, having completed my first academic year with an impeccable GPA of 4.0. With a solid academic foundation and my previous experience as a grader in my undergraduate, I am confident in my ability to contribute effectively to the course's grading process. Additionallymy two years of experience as a software developer and have hands-on experience in the field. I am excited about the prospect of leveraging my knowledge and experience to provide valuable assistance in grading assignments, tests, and projects. I eagerly await your response regarding the availability of this opportunity and the chance to further discuss my qualifications. 

I have enclosed my resume, cover letter, and transcripts for your review. Thank you for your time and consideration. 
    
    
Best Regards,
Nitej
+1 (602)784-5356"""

    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = subject
    body = MIMEText(content, 'plain')
    msg.attach(body)
    for f in ["C:/Users/mmithint/Downloads/transcript_ms.pdf","C:/Users/mmithint/Downloads/Nitej_G.pdf", "C:/Users/mmithint/Desktop/resume/GRADESHEETB.pdf"]:
            with open(f, "rb") as fil:
                part = MIMEApplication(
                    fil.read(),
                    Name=basename(f)
                )
            # After the file is closed
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
            msg.attach(part)
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(from_addr, 'aooiuzwqznvcwlbn')
        # 'fxilahdzpkcgfyef'
        server.send_message(msg, from_addr=from_addr, to_addrs=[to_addr])
        print(f"email sent {SNO}")
    except Exception as e:
         print(e)

data = pd.read_excel("C:\\Users\\mmithint\\Downloads\\graderdatasheet.xlsx")

for i in data.index:
    print(data["ProffesorName"][i], data["ProffesorEmail"][i], data["CourseNumber"][i], data["CourseName"][i])
    MailToProffesor(data["SNO"], data["ProffesorName"][i], data["ProffesorEmail"][i], data["CourseNumber"][i], data["CourseName"][i])
