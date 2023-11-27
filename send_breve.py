import pandas as pd
import numpy as np
from redmail import outlook


def get_secretaries():
    file = pd.read_excel("Oversigt.xlsx")
    secretaries = {"Institut for Matematik og Computer Science": file["Mailkontakt"][0],
                   "Institut for Fysik": file["Mailkontakt"][1],
                   "Institut for Miljø og Ressourceteknologi": file["Mailkontakt"][2],
                   "Institut for Sundhedsteknologi": file["Mailkontakt"][3],
                   "Institut for Fødevareinstitut": file["Mailkontakt"][4],
                   "Institut for Akvatiske Ressourcer": file["Mailkontakt"][5],
                   "Institut for Kemi": file["Mailkontakt"][6],
                   "Institut for Bioteknologi og Biomedicin": file["Mailkontakt"][7],
                   "Institut for Kemiteknik": file["Mailkontakt"][8],
                   "Novo Nordisk Foundation Center for Biosustainability": file["Mailkontakt"][9],
                   "Institut for Rumforskning og Rumteknologi": file["Mailkontakt"][10],
                   "Institut for Elektroteknologi og Fotonik": file["Mailkontakt"][11],
                   "Institut for Byggeri og Mekanisk Teknologi": file["Mailkontakt"][12],
                   "Institut for Teknologi, Ledelse og Økonomi": file["Mailkontakt"][13],
                   "Institut for Vindenergi og Energisystemer": file["Mailkontakt"][14],
                   "Institut for Energikonvertering og -lagring": file["Mailkontakt"][15],
                   "Nationalt Center for Nanofabrikation og -karakterisering": file["Mailkontakt"][16],
                   "Danish Offshore Technology Centre": file["Mailkontakt"][17],
                   }
    return secretaries


def get_students():
    file = pd.read_excel("Oktober.xlsx")
    file = file.dropna()
    students = np.array([file["Navn"], file["Institut"]])
    return students


# def send_email():
#     body = "Subject: Subject here . \n" \
#            "Dear ContactName, \n\n" + \
#            "Email Body Here \n This is a new line. \n\n This is a double new line" + \
#            "\nBest regards \n David"
#     try:
#         smtpObj = smtplib.SMTP('smtp.office365.com', 587)
#     except Exception as e:
#         print(e)
#         smtpObj = smtplib.SMTP_SSL('smtp.office365.com', 465)
#
#     smtpObj.ehlo()
#     smtpObj.starttls()
#     smtpObj.login('davos@dtu.dk', 'DTUBerta1998')
#     smtpObj.sendmail('davos@dtu.dk', 'davidostenfeldt@hotmail.com', body)
#
#     smtpObj.quit()
#     pass

def send_email():
    outlook.username = 'testmail@test.com'
    outlook.password = "123456789"
    outlook.send(
        receivers=['s194237@student.dtu.dk'],
        subject='A test subject',
        text='Dear David, <br> This is the body of the text <br> Best regards, <br> David'
    )

def main():
    secretaries = get_secretaries()
    students = get_students()

    send_email()

    for i in range(len(students[0])):
        to = students[0][i]
        cc = secretaries[students[1][i]]
        print("to: ", to, "cc: ", cc)


if __name__ == '__main__':
    main()
