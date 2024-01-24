import pandas as pd
import numpy as np
import win32com.client as win32
import os


def get_secretaries(file_path, brev_type='4mdr'):
    file = pd.read_excel(file_path)
    if brev_type == '4mdr' or brev_type == '4-mdrs. brev':
        secretaries = {"Institut for Matematik og Computer Science": file["Mailkontakt"][0],
                    "Institut for Fysik": file["Mailkontakt"][1],
                    "Institut for Miljø og Ressourceteknologi": file["Mailkontakt"][2],
                    "Institut for Sundhedsteknologi": file["Mailkontakt"][3],
                    "Fødevareinstituttet": file["Mailkontakt"][4],
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
    else:
        secretaries = {"DTU Compute": file["Mailkontakt"][0],
                    "DTU Fysik": file["Mailkontakt"][1],
                    "DTU Sustain": file["Mailkontakt"][2],
                    "DTU Sundhedsteknologi": file["Mailkontakt"][3],
                    "DTU Sund": file["Mailkontakt"][3],
                    "DTU Fødevareinstituttet": file["Mailkontakt"][4],
                    "DTU Aqua": file["Mailkontakt"][5],
                    "DTU Kemi": file["Mailkontakt"][6],
                    "DTU Bioengineering": file["Mailkontakt"][7],
                    "DTU Kemiteknik": file["Mailkontakt"][8],
                    "DTU Biosustain": file["Mailkontakt"][9],
                    "DTU Space": file["Mailkontakt"][10],
                    "DTU Electro": file["Mailkontakt"][11],
                    "DTU Construct": file["Mailkontakt"][12],
                    "DTU Management": file["Mailkontakt"][13],
                    "DTU Wind": file["Mailkontakt"][14],
                    "DTU Energi": file["Mailkontakt"][15],
                    "DTU Nanolab": file["Mailkontakt"][16],
                    "DTU Offshore": file["Mailkontakt"][17],
                    }
    return secretaries


def get_students(file_path):
    file = pd.read_excel(file_path)
    file = file.dropna()
    students = np.array([file["Navn"], file["Institut"]])
    return students

def get_vedh_filer(folder_path):
    vedh_filer = []
    # Get path for all files in folder
    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            vedh_filer.append(os.path.join(folder_path, file))
    
    # Sort files by number in name
    vedh_filer.sort(key=lambda f: int(''.join(filter(str.isdigit, f))))

    return vedh_filer


def send_email(recipient_name, cc_name, subject_text, body_text, attachment=''):
    outlook = win32.Dispatch('outlook.application')
    namespace = outlook.GetNamespace('MAPI')
    mail = outlook.CreateItem(0)
    mail.Subject = subject_text
    mail.Body = body_text
    recipient = mail.Recipients.Add(recipient_name)
    recipient.Resolve()
    cc_recipient = mail.Recipients.Add(cc_name)
    cc_recipient.Type = 2 # 2 corresponds to CC
    cc_recipient.Resolve()

    if recipient.Resolved and cc_recipient.Resolved:
        if attachment != '':
            mail.Attachments.Add(attachment)
        mail.Send()
    else:
        if not recipient.Resolved:
            print(f'Could not resolve recipient: {recipient_name}')
        if not cc_recipient.Resolved:
            print(f'Could not resolve CC recipient: {cc_name}')

def main():
    secretaries = get_secretaries(r'C:\Users\s194237\OneDrive - Danmarks Tekniske Universitet\Skrivebord\Oversigt over skoleleder og sekretær.xlsx', type="4mdr")
    students = get_students(r'C:\Users\s194237\OneDrive - Danmarks Tekniske Universitet\Skrivebord\Marts 2024.xlsx')

    send_email('David Ari Ostenfeldt', 'David Ari Ostenfeldt', 'This is the subject', 'This is the body \nWith newlines \nwoooow\n\nBest regards,\nDavid Ari Ostenfeldt')

    # for i in range(len(students[0])):
    #     to = students[0][i]
    #     cc = secretaries[students[1][i]]
    #     print("to: ", to, "cc: ", cc)


if __name__ == '__main__':
    main()
