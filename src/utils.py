import pandas as pd
import numpy as np
import win32com.client as win32
import os

def get_secretaries(brev_type='4mdr', file_path='Oversigt over skoleleder og sekretær.xlsx', ):
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


def get_vejledere(file_path='Vejledere.xlsx'):
    file = pd.read_excel(file_path)
    file = file.dropna()
    # TODO: FIX PATH
    vejledere = np.array([file["Hovedvejleder"], file["Institut"]])
    return vejledere


def get_vedh_filer(folder_path):
    vedh_filer = []
    # Get path for all files in folder
    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            vedh_filer.append(os.path.join(folder_path, file))
    
    # Sort files by number in name
    vedh_filer.sort(key=lambda f: int(''.join(filter(str.isdigit, f))))

    return vedh_filer


def get_final_recommendations(folder_path):
    final_recommendations = []
    # Get path for all files in folder
    for file in os.listdir(folder_path):
        if file.endswith(".pdf") or file.endswith(".docx"):
            final_recommendations.append(os.path.join(folder_path, file))
    
    # Sort files by number in name
    final_recommendations.sort(key=lambda f: int(''.join(filter(str.isdigit, f))))

    return final_recommendations


def get_best_matching_final_recommendation(student_name, final_recommendations):
    best_match_score = 0
    best_recommendation = ""
    for recommendation in final_recommendations:
        score = 0
        for name in student_name.split(' '):
            if name.lower() in recommendation.lower():
                score += 1
        if score > best_match_score:
            best_match_score = score
            best_recommendation = recommendation

    return best_recommendation


def test_recipient(recipient_name):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    recipient = mail.Recipients.Add(recipient_name)
    recipient.Resolve()
    return recipient.Resolved


def send_email(recipient_name, cc_name, subject_text, body_text, attachment=''):
    outlook = win32.Dispatch('outlook.application')
    #namespace = outlook.GetNamespace('MAPI')
    mail = outlook.CreateItem(0)
    mail.Subject = subject_text
    mail.Body = body_text
    recipient = mail.Recipients.Add(recipient_name)
    recipient.Resolve()
    cc_recipient = mail.Recipients.Add(cc_name)
    cc_recipient.Type = 2 # 2 corresponds to CC
    cc_recipient.Resolve()

    unresolved_cc = ""
    unresolved_recipient = ""
    if recipient.Resolved and cc_recipient.Resolved:
        if attachment != '':
            for path in attachment.split(';'):
                mail.Attachments.Add(path)
            # mail.Attachments.Add(attachment)
        mail.Send()
    else:
        if not recipient.Resolved:
            print(f'Could not resolve recipient: {recipient_name}')
            unresolved_recipient = recipient_name
        if not cc_recipient.Resolved:
            print(f'Could not resolve CC recipient: {cc_name}')
            unresolved_cc = cc_name

    return unresolved_recipient, unresolved_cc


def display_email(recipient_name, cc_name, subject_text, body_text, attachment=''):
    outlook = win32.Dispatch('outlook.application')
    #namespace = outlook.GetNamespace('MAPI')
    mail = outlook.CreateItem(0)
    mail.Subject = subject_text
    mail.Body = body_text
    recipient = mail.Recipients.Add(recipient_name)
    recipient.Resolve()
    cc_recipient = mail.Recipients.Add(cc_name)
    cc_recipient.Type = 2 # 2 corresponds to CC
    cc_recipient.Resolve()

    unresolved_recipient = ""
    unresolved_cc = ""
    if recipient.Resolved and cc_recipient.Resolved:
        if attachment != '':
            for path in attachment.split(';'):
                mail.Attachments.Add(path)
            #mail.Attachments.Add(attachment)
        mail.display()
    else:
        if not recipient.Resolved:
            print(f'Could not resolve recipient: {recipient_name}')
            unresolved_recipient = recipient_name
        if not cc_recipient.Resolved:
            print(f'Could not resolve CC recipient: {cc_name}')
            unresolved_cc = cc_name

    return unresolved_recipient, unresolved_cc


