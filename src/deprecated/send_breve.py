import pandas as pd
import numpy as np
import win32com.client as win32
import os
import enchant



def get_final_recommendations(folder_path):
    final_recommendations = []
    # Get path for all files in folder
    for file in os.listdir(folder_path):
        if file.endswith(".pdf") or file.endswith(".docx"):
            final_recommendations.append(os.path.join(folder_path, file))
    
    # Sort files by number in name
    final_recommendations.sort(key=lambda f: int(''.join(filter(str.isdigit, f))))

    return final_recommendations



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

    unresolved = []
    if recipient.Resolved and cc_recipient.Resolved:
        if attachment != '':
            for path in attachment.split(';'):
                mail.Attachments.Add(path)
            # mail.Attachments.Add(attachment)
        mail.Send()
    else:
        if not recipient.Resolved:
            print(f'Could not resolve recipient: {recipient_name}')
            unresolved += [recipient_name]
        if not cc_recipient.Resolved:
            print(f'Could not resolve CC recipient: {cc_name}')
            unresolved += [cc_name]

    return unresolved


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

    unresolved = []
    if recipient.Resolved and cc_recipient.Resolved:
        if attachment != '':
            for path in attachment.split(';'):
                mail.Attachments.Add(path)
            #mail.Attachments.Add(attachment)
        mail.display()
    else:
        if not recipient.Resolved:
            print(f'Could not resolve recipient: {recipient_name}')
            unresolved += [recipient_name]
        if not cc_recipient.Resolved:
            print(f'Could not resolve CC recipient: {cc_name}')
            unresolved += [cc_name]

    return unresolved

def main():
    #secretaries = get_secretaries(r'C:\Users\s194237\OneDrive - Danmarks Tekniske Universitet\Skrivebord\Oversigt over skoleleder og sekretær.xlsx', type="4mdr")
    #students = get_students(r'C:\Users\s194237\OneDrive - Danmarks Tekniske Universitet\Skrivebord\Marts 2024.xlsx')

    send_email('David Ari Ostenfeldt', 'David Ari Ostenfeldt', 'This is the subject', 'This is the body \nWith newlines \nwoooow\n\nBest regards,\nDavid Ari Ostenfeldt')

    # for i in range(len(students[0])):
    #     to = students[0][i]
    #     cc = secretaries[students[1][i]]
    #     print("to: ", to, "cc: ", cc)


if __name__ == '__main__':
    main()
