import tkinter as tk
from typing import Tuple
import customtkinter as ctk
import pandas as pd
import os
import utils



class MainApplication(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        # Configure the main window
        self.title('BrevGUI')
        self.geometry(f'{650}x{625}')

        # Configure the layout of the main window
        self.grid_columnconfigure((0,1), weight=1)
        self.grid_rowconfigure((0,1,2,3,5), weight=1)
        self.grid_rowconfigure(4, weight=3)

        ##### ROW 1 #####
        self.letter_type_frame = LetterTypeFrame(self, width=300, corner_radius=0)
        self.letter_type_frame.grid(row=0, column=0, sticky='nsew')
        self.letter_type_frame.grid_rowconfigure(0, weight=1)
        self.letter_type_frame.grid_rowconfigure(1, weight=2)
        
        self.sender_frame = SenderFrame(self, width=300, corner_radius=0)
        self.sender_frame.grid(row=0, column=1, sticky='nsew')
        self.sender_frame.grid_rowconfigure(0, weight=1)
        self.sender_frame.grid_rowconfigure(1, weight=2)

        ##### ROW 2 #####
        self.file_frame = FileFrame(self, width=300, corner_radius=0)
        self.file_frame.grid(row=1, column=0, sticky='nsew')
        self.file_frame.grid_rowconfigure(0, weight=1)
        self.file_frame.grid_rowconfigure(1, weight=2)

        self.recipient_frame = RecipientFrame(self, width=300, corner_radius=0)
        self.recipient_frame.grid(row=1, column=1, sticky='nsew')
        self.recipient_frame.grid_rowconfigure(0, weight=1)
        self.recipient_frame.grid_rowconfigure(1, weight=2)

        ##### ROW 3 #####
        self.language_frame = LanguageFrame(self, width=300, corner_radius=0)
        self.language_frame.grid(row=2, column=0, columnspan=2, sticky='nsew')
        self.language_frame.grid_rowconfigure(0, weight=1)
        self.language_frame.grid_columnconfigure(0, weight=1)
        self.language_frame.grid_columnconfigure((1,2,3), weight=3)

        ##### ROW 4 #####
        self.subject_frame = SubjectFrame(self, width=600, corner_radius=0)
        self.subject_frame.grid(row=3, column=0, columnspan=2, sticky='nsew')
        self.subject_frame.grid_rowconfigure(0, weight=1)
        self.subject_frame.grid_columnconfigure(0, weight=1)
        self.subject_frame.grid_columnconfigure(1, weight=3)

        ##### ROW 5 #####
        self.content_frame = ContentFrame(self, width=600, corner_radius=0)
        self.content_frame.grid(row=4, column=0, columnspan=2, sticky='nsew')
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=3)
        self.content_frame.grid_columnconfigure(0, weight=1)

        ##### ROW 6 #####
        self.send_button_frame = BottomFrame(self, width=600, corner_radius=0)
        self.send_button_frame.grid(row=5, column=0, columnspan=2, sticky='nsew')
        self.send_button_frame.grid_rowconfigure(0, weight=1)
        self.send_button_frame.grid_columnconfigure((0,1,2), weight=1)
        
        # Setup popup window
        self.popup_window = None
        self.popup_window_send = None
        
        self.update_send_button()
     
        
    def set_content(self, _current_value=None):
        brev_type = self.letter_type_frame.get()

        if brev_type not in self.letter_type_frame.brevtyper:
            self.open_popup_window(size=[300, 125], title="Fejl", message="Brevtype er ikke valgt eller ugyldig.")
            return
        
        if self.language_frame.get() == 'auto':
            _language = 'da'
        else:
            _language = self.language_frame.get()

        text_file = open('content/' + _language + '-' + brev_type + '.txt', 'r', encoding='utf-8')
        text = text_file.read()
        text_file.close()

        # Set subject
        subject = text.split('Emne: ')[1].split('\n')[0]
        self.subject_frame.set(subject)

        # Set content
        content = text.split('Indhold:\n')[1]
        self.content_frame.set(content)
        
    
    def open_popup_window(self, size=[300, 125], title="Popup window", message="This is a popup window."):
        if self.popup_window is None or not self.popup_window.winfo_exists():
            self.popup_window = PopupWindow(self, size, title, message)
            self.popup_window.focus()
        else:
            self.popup_window.focus()

    def open_send_popup_window(self, size=[200, 125], unresolved_names=[]):
        if self.popup_window_send is None or not self.popup_window_send.winfo_exists():
            self.popup_window_send = SendPopupWindow(self, size, unresolved_names)
            self.popup_window_send.focus()
        else:
            self.popup_window_send.focus()
            

    def update_send_button(self):
        # Do a manual check, to not incur errors
        if isinstance(self.file_frame, FileFrame):
            if self.letter_type_frame.brevtype_var.get() == "Vælg brevtype" or self.file_frame.file_button._text == 'Vælg mappe med breve' or self.recipient_frame.recipient_button._text == 'Vælg modtager liste':
                self.send_button_frame.send_button.configure(state='disabled')
            else:
                self.send_button_frame.send_button.configure(state='normal')
        else:
            if self.letter_type_frame.brevtype_var.get() == "Vælg brevtype" or self.file_frame.letters_button._text == 'Vælg mappe med breve' or self.file_frame.final_recommendations_button._text == 'Vælg mappe med\nfinal recommendations' or self.recipient_frame.recipient_button._text == 'Vælg modtager liste':
                self.send_button_frame.send_button.configure(state='disabled')
            else:
                self.send_button_frame.send_button.configure(state='normal')

    
    def test_breve(self):
        if self.sender_frame.get() is None:
            self.open_popup_window(size=[300, 125], title="Fejl", message="Afsender navn ikke indtastet.")
            return
            
        
        self.letter_type = self.letter_type_frame.get()
        if self.letter_type == '4-mdrs. brev' or self.letter_type == 'Gradbrev (stud.)':
            recipients = utils.get_students(self.recipient_frame.get())
        elif self.letter_type == 'Gradbrev (vejl.)':
            recipients = utils.get_vejledere(self.recipient_frame.get())
            students = utils.get_students(self.recipient_frame.get())
            
        secretaries = utils.get_secretaries(self.letter_type_frame.get())
        
        # Test secretaries
        for i in range(len(recipients[1])):
            try:
                cc = secretaries[recipients[1][i]]
            except:
                self.open_popup_window(size=[300, 125], title="Fejl", message=f"Kunne ikke finde sekretær for {recipients[1][i]} i modtagerliste.\nTjek om du har valgt den rigtige brevtype og modtagerliste.")
                return

        final_recommendations_path = None
        if isinstance(self.file_frame, FileFrame):
            vedh_filer_path = self.file_frame.get()
        else:
            vedh_filer_path = self.file_frame.get(self.file_frame.letters_button)
            final_recommendations_path = self.file_frame.get(self.file_frame.final_recommendations_button)

        vedh_filer = utils.get_vedh_filer(vedh_filer_path)
        if len(vedh_filer) != len(recipients[0]):
            self.open_popup_window(size=[300, 125], title="Fejl", message="Antal vedhæftede filer stemmer ikke overens med antal modtagere.")
            return

        if final_recommendations_path is not None:
            final_recommendations = utils.get_final_recommendations(final_recommendations_path)
            if len(final_recommendations) != len(recipients[0]):
                self.open_popup_window(size=[300, 125], title="Fejl", message="Antal final recommendations stemmer ikke overens med antal modtagere.")
                return
        

        # Test recipients
        resolved_recipients = []
        unresolved_recipients = []
        resolved_secretaries = []
        unresolved_secretaries = []
        resolved_vedh_filer = []
        unresolved_vedh_filer = []
        if self.letter_type == 'Gradbrev (vejl.)':
            resolved_students = []
            unresolved_students = []

        for i in range(len(recipients[0])):
            if not utils.test_recipient(recipients[0][i]): # If we cannot find the recipient in the system
                unresolved_recipients.append(recipients[0][i])
                unresolved_secretaries.append(secretaries[recipients[1][i]])
                unresolved_vedh_filer.append(vedh_filer[i])
                if self.letter_type == 'Gradbrev (vejl.)':
                    unresolved_students.append(students[0][i])
            else:
                resolved_recipients.append(recipients[0][i])
                resolved_secretaries.append(secretaries[recipients[1][i]])
                resolved_vedh_filer.append(vedh_filer[i])
                if self.letter_type == 'Gradbrev (vejl.)':
                    resolved_students.append(students[0][i])

        # Set all necessary variables
        self.resolved_recipients = resolved_recipients
        self.unresolved_recipients = unresolved_recipients
        self.resolved_secretaries = resolved_secretaries
        self.unresolved_secretaries = unresolved_secretaries
        self.resolved_vedh_filer = resolved_vedh_filer
        self.unresolved_vedh_filer = unresolved_vedh_filer
        if self.letter_type == 'Gradbrev (stud.)':
            self.final_recommendations = final_recommendations
        if self.letter_type == 'Gradbrev (vejl.)':
            self.resolved_students = resolved_students
            self.unresolved_students = unresolved_students

        # Open popup window with unresolved names
        if len(self.unresolved_recipients) > 0:
            self.open_send_popup_window(unresolved_names=self.unresolved_recipients)
        else:
            self.send_breve()
        


    def send_breve(self):
        # TODO: Implement auto detecting language (see set_content() method)
        num_recipients = len(self.resolved_recipients)
        
        for i in range(num_recipients):
            recipient = self.resolved_recipients[i]
            vedh_fil = self.resolved_vedh_filer[i]
            cc = self.resolved_secretaries[i]

            subject_text = self.subject_frame.get()
            content_text = self.content_frame.get()
            
            # Replace placeholder text
            self.letter_type = self.letter_type_frame.get()
            if self.letter_type == 'Gradbrev (vejl.)': 
                student = self.resolved_students[i]
                subject_text = subject_text.replace('[student navn]', student)
                content_text = content_text.replace('[student navn]', student)
            
            content_text = content_text.replace('[modtager navn]', recipient)
            content_text = content_text.replace('[afsender navn]', self.sender_frame.get())

            if self.letter_type_frame.get() == 'Gradbrev (stud.)':
                best_final_recommendation = utils.get_best_matching_final_recommendation(recipient, self.final_recommendations)
                if best_final_recommendation == "":
                    self.open_popup_window(size=[300, 125], title="Fejl", message=f"Kunne ikke finde final recommendation for {recipient}.")
                    return
                vedh_fil = vedh_fil + ';' + best_final_recommendation
            

            unresolved_list = []
            unresolved_cc_list = []
            self.send_mode = self.send_button_frame.get_send_mode()
            if self.send_mode == 'auto':
                unresolved_recipient, unresolved_cc = utils.display_email(recipient, cc, subject_text, content_text, vedh_fil)
                if unresolved_recipient != "":
                    unresolved_list.append(unresolved_recipient)
                if unresolved_cc != "":
                    unresolved_cc_list.append(unresolved_cc)
            elif self.send_mode == 'manuel':
                unresolved_recipient, unresolved_cc = utils.display_email(recipient, cc, subject_text, content_text, vedh_fil)
                if unresolved_recipient != "":
                    unresolved_list.append(unresolved_recipient)
                if unresolved_cc != "":
                    unresolved_cc_list.append(unresolved_cc)

        if len(unresolved_list) > 0:
            self.open_popup_window(size=[300, 125], title="Advarsel!", message=f"Breve sendt til {num_recipients - len(unresolved_list)}\nKunne ikke sende breve til {len(unresolved_list)} modtagere:\n{', '.join(unresolved_list)}")
        elif len(unresolved_cc_list) > 0:
            self.open_popup_window(size=[300, 125], title="Advarsel!", message=f"Breve sendt til {num_recipients - len(unresolved_cc_list)}\nKunne ikke sende breve til {len(unresolved_cc_list)} modtagere:\n{', '.join(unresolved_cc_list)}")
        else:
            self.open_popup_window(size=[300, 125], title=f"Succes: {num_recipients} breve sendt", message=f"Der blev sendt breve til alle {num_recipients} modtagere.")

            

class LetterTypeFrame(ctk.CTkFrame):
    def __init__(self, master, width=300, corner_radius=0):
        super().__init__(master, width, corner_radius)
        self.master = master
        self.width = width
        self.corner_radius = corner_radius
        
        # Label
        self.brevtype_label = ctk.CTkLabel(self, text='Brevtype', font=ctk.CTkFont(size=20, weight='bold'))
        self.brevtype_label.grid(row=0, column=0, padx=10, pady=(10,5), sticky='nsew')

        # OptionMenu
        self.brevtyper = ['4-mdrs. brev', 'Gradbrev (stud.)', 'Gradbrev (vejl.)']
        self.brevtype_var = tk.StringVar()
        self.brevtype_option_menu = ctk.CTkOptionMenu(self, variable=self.brevtype_var, command=self.set, values=self.brevtyper, width=275)
        self.brevtype_option_menu.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')
        self.brevtype_option_menu.set('Vælg brevtype')


    def get(self):
        return self.brevtype_var.get()
    

    def set(self, value):
        if self.master.file_frame.winfo_exists():
            if value == 'Gradbrev (stud.)' and isinstance(self.master.file_frame, FileFrame):
                self.master.file_frame.destroy()
                self.master.file_frame = MultipleFilesFrame(self.master, width=300, corner_radius=0)
            elif isinstance(self.master.file_frame, MultipleFilesFrame):
                self.master.file_frame.destroy()
                self.master.file_frame = FileFrame(self.master, width=300, corner_radius=0)
            
            self.master.file_frame.grid(row=1, column=0, sticky='nsew')
            self.master.file_frame.grid_rowconfigure(0, weight=1)
            self.master.file_frame.grid_rowconfigure(1, weight=2)
                
        self.master.set_content()
        self.master.update_send_button()


class SenderFrame(ctk.CTkFrame):
    def __init__(self, master, width=300, corner_radius=0):
        super().__init__(master, width, corner_radius)
        self.master = master
        self.width = width
        self.corner_radius = corner_radius
        
        # Label
        self.sender_label = ctk.CTkLabel(self, text='Afsender', font=ctk.CTkFont(size=20, weight='bold'))
        self.sender_label.grid(row=0, column=0, padx=10, pady=(10,5))

        # Entry
        self.sender_entry = ctk.CTkEntry(self, placeholder_text='Afsender', width=275)
        self.sender_entry.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')
        
    def get(self):
        text = self.sender_entry.get()
        if text == '':
            self.master.open_popup_window(size=[300, 125], title="Fejl", message="Ingen afsender valgt.")
            return None
        return text


class FileFrame(ctk.CTkFrame):
    def __init__(self, master, width=300, corner_radius=0):
        super().__init__(master, width, corner_radius)
        self.master = master
        self.width = width
        self.corner_radius = corner_radius
        
        # Label
        self.file_label = ctk.CTkLabel(self, text='Vedhæftede filer', font=ctk.CTkFont(size=20, weight='bold'))
        self.file_label.grid(row=0, column=0, padx=10, pady=(10,5))

        # Button
        self.file_button = ctk.CTkButton(self, text='Vælg mappe med breve', width=275, command=self.browse_files)
        self.file_button.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')

        # Initialize filename
        self.filename = None
        
    def browse_files(self):
        self.filename =  ctk.filedialog.askdirectory(
            initialdir='/',
            title="Vælg mappe med breve",
            parent=self.master
        )
        
        if self.filename != "":
            self.set(self.filename)
        else: 
            self.set('Vælg mappe med breve')
        
    
    def get(self):
        if self.filename is None:
            self.master.open_popup_window(size=[300, 125], title="Fejl", message="Ingen mappe valgt.")
        else:
            return self.filename
        
    def set(self, value):
        if value != 'Vælg mappe med breve':
            self.file_button.configure(text=value.split('/')[-1])
        else:
            self.file_button.configure(text=value)
        self.filename = value
        self.master.update_send_button()


class MultipleFilesFrame(ctk.CTkFrame):
    def __init__(self, master, width=300, corner_radius=0):
        super().__init__(master, width, corner_radius)
        self.master = master
        self.width = width
        self.corner_radius = corner_radius

        # Label
        self.file_label = ctk.CTkLabel(self, text='Vedhæftede filer', font=ctk.CTkFont(size=20, weight='bold'))
        self.file_label.grid(row=0, column=0, padx=10, pady=(10,5), columnspan=2, sticky='nsew')

        # Button 1
        self.letters_button = ctk.CTkButton(self, text='Vælg mappe med breve', width=140, command=lambda: self.browse_files(self.letters_button))
        self.letters_button.grid(row=1, column=0, padx=(10,5), pady=(0,10), sticky='nsew')

        # Button 2
        self.final_recommendations_button = ctk.CTkButton(self, text='Vælg mappe med\nfinal recommendations', width=140, command=lambda: self.browse_files(self.final_recommendations_button))
        self.final_recommendations_button.grid(row=1, column=1, padx=(5,10), pady=(0,10), sticky='nsew')

        # Initialize filenames
        self.letters_filename = None
        self.final_recommendations_filename = None

    def browse_files(self, button):
        if button == self.letters_button:
            self.letters_filename =  ctk.filedialog.askdirectory(
                initialdir='/',
                title="Vælg mappe med breve",
                parent=self.master
            )

            if self.letters_filename != "":
                self.set(self.letters_filename, button)
            else:
                self.set('Vælg mappe med breve', button)

        elif button == self.final_recommendations_button:
            self.final_recommendations_filename =  ctk.filedialog.askdirectory(
                initialdir='/',
                title="Vælg mappe med final recommendations",
                parent=self.master
            )

            if self.final_recommendations_filename != "":
                self.set(self.final_recommendations_filename, button)
            else:
                self.set('Vælg mappe med\nfinal recommendations', button)

    
    def get(self, button):
        if button == self.letters_button:
            if self.letters_filename is None:
                self.master.open_popup_window(size=[300, 125], title="Fejl", message="Ingen mappe med breve valgt.")
                return
            return self.letters_filename
        elif button == self.final_recommendations_button:
            if self.final_recommendations_filename is None:
                self.master.open_popup_window(size=[300, 125], title="Fejl", message="Ingen mappe med final recommendations valgt.")
                return
            return self.final_recommendations_filename


    def set(self, value, button):
        if button == self.letters_button:
            if value != 'Vælg mappe med breve':
                self.letters_button.configure(text=value.split('/')[-1])
            else:
                self.letters_button.configure(text=value)
            self.letters_filename = value
        elif button == self.final_recommendations_button:
            if value != 'Vælg mappe med final recommendations':
                self.final_recommendations_button.configure(text=value.split('/')[-1])
            else:
                self.final_recommendations_button.configure(text=value)
            self.final_recommendations_filename = value
        self.master.update_send_button()
        


class RecipientFrame(ctk.CTkFrame):
    def __init__(self, master, width=300, corner_radius=0):
        super().__init__(master, width, corner_radius)
        self.master = master
        self.width = width
        self.corner_radius = corner_radius
        
        # Label
        self.recipient_label = ctk.CTkLabel(self, text='Vælg modtager liste', font=ctk.CTkFont(size=20, weight='bold'))
        self.recipient_label.grid(row=0, column=0, padx=10, pady=(10,5))

        # Button
        self.recipient_button = ctk.CTkButton(self, text='Vælg modtager liste', width=275, command=self.browse_files)
        self.recipient_button.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')
        
        # Initialize filename
        self.filename = None
        
    def get(self):
        if self.filename == None:
            self.master.open_popup_window(size=[300, 125], title="Fejl", message="Ingen modtagerliste valgt.")
            return
        return self.filename
    
    def set(self, value):
        if value != 'Vælg modtager liste':
            self.recipient_button.configure(text=value.split('/')[-1])
        else:
            self.recipient_button.configure(text=value)
        self.filename = value
        self.master.update_send_button()
        
    def browse_files(self):
        self.filename =  ctk.filedialog.askopenfilename(
            initialdir='/',
            title="Vælg modtager liste",
            filetypes=[("Excel files", "*.xlsx")],
            parent=self.master
        )
        
        if self.filename != "":
            self.set(self.filename)
        else:
            self.set('Vælg modtager liste')


class LanguageFrame(ctk.CTkFrame):
    def __init__(self, master, width=300, corner_radius=0):
        super().__init__(master, width, corner_radius)
        self.master = master
        self.width = width
        self.corner_radius = corner_radius
    
        # Label
        self.language_label = ctk.CTkLabel(self, text='Sprog:', font=ctk.CTkFont(size=20, weight='bold'))
        self.language_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsw')

        # Radiobuttons
        self.language_var = tk.StringVar()
        self.language_var.set('da')
        self.language_radiobutton1 = ctk.CTkRadioButton(self, text='Dansk', command=self.master.set_content, variable=self.language_var, value='da')
        self.language_radiobutton1.grid(row=0, column=1, padx=(10,5), pady=10, sticky='nse')
        self.language_radiobutton2 = ctk.CTkRadioButton(self, text='Engelsk', command=self.master.set_content, variable=self.language_var, value='en')
        self.language_radiobutton2.grid(row=0, column=2, padx=0, pady=10, sticky='nse')
        self.language_radiobutton3 = ctk.CTkRadioButton(self, text='Automatisk', command=self.master.set_content, variable=self.language_var, value='auto')
        self.language_radiobutton3.grid(row=0, column=3, padx=(5,10), pady=10, sticky='nse')

    def get(self):
        return self.language_var.get()


class SubjectFrame(ctk.CTkFrame):
    def __init__(self, master, width=600, corner_radius=0):
        super().__init__(master, width, corner_radius)
        self.master = master
        self.width = width
        self.corner_radius = corner_radius
        
        # Label
        self.subject_label = ctk.CTkLabel(self, text='Emne:', font=ctk.CTkFont(size=20, weight='bold'))
        self.subject_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsw')

        # Entry
        self.subject_entry = ctk.CTkEntry(self, placeholder_text='Emne', width=500)
        self.subject_entry.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')
        
    def get(self):
        return self.subject_entry.get()
    
    def set(self, value):
        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, value)


class ContentFrame(ctk.CTkFrame):
    def __init__(self, master, width=600, corner_radius=0):
        super().__init__(master, width, corner_radius)
        self.master = master
        self.width = width
        self.corner_radius = corner_radius
        
        # Label
        self.content_label = ctk.CTkLabel(self, text='Indhold tekst:', font=ctk.CTkFont(size=20, weight='bold'))
        self.content_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsw')

        # Text
        self.content_text = ctk.CTkTextbox(self)
        self.content_text.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')
        
    def get(self):
        return self.content_text.get(1.0, tk.END)
    
    def set(self, value):
        self.content_text.delete(1.0, tk.END)
        self.content_text.insert(tk.END, value)


class BottomFrame(ctk.CTkFrame):
    def __init__(self, master, width=600, corner_radius=0):
        super().__init__(master, width, corner_radius)
        self.master = master
        self.width = width
        self.corner_radius = corner_radius
        
        # See recipients button
        self.see_recipients_button = ctk.CTkButton(self, text='Se modtagere', width=250, height=50, command=self.open_recipients_window)
        self.see_recipients_button.grid(row=0, column=0, rowspan=2, padx=10, pady=10, sticky='nsw')

        # Send mode radio buttons
        self.send_mode_var = tk.StringVar()
        self.send_mode_var.set('auto')
        self.send_mode_radiobutton1 = ctk.CTkRadioButton(self, text='Auto', variable=self.send_mode_var, value='auto')
        self.send_mode_radiobutton1.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')
        self.send_mode_radiobutton2 = ctk.CTkRadioButton(self, text='Manuel', variable=self.send_mode_var, value='manuel')
        self.send_mode_radiobutton2.grid(row=1, column=1, padx=10, pady=10, sticky='nsew')

        # Send Button
        self.send_button = ctk.CTkButton(self, text='Send breve', width=250, height=50, command=self.send)
        self.send_button.grid(row=0, column=2, rowspan=2, padx=10, pady=10, sticky='nse')

        # Set recipients window
        self.recipients_window = None
        
    def get_send_mode(self):
        return self.send_mode_var.get()
    
    def open_recipients_window(self):
        if self.recipients_window is None or not self.recipients_window.winfo_exists():
            self.recipients_window = RecipientsWindow(self, self.master.recipient_frame.get())
            if self.recipients_window.winfo_exists():
                self.recipients_window.focus()
        else:
            self.recipients_window.focus()
    
    def send(self):
        # It is important that the test_breve method is called before the send_breve method
        # DO NOT call send_breve directly, it will be handled by test_breve
        self.master.test_breve()


class RecipientsWindow(ctk.CTkToplevel):
    def __init__(self, root, recipient_path, size=[400, 400]) -> None:
        super().__init__()
        
        self.root = root
        self.main_application = root.master
        self.size = size
        self.recipient_path = recipient_path
        self.letter_type = self.main_application.letter_type_frame.get()
        
        if self.recipient_path is None:
            self.main_application.open_popup_window(size=[300, 125], title="Fejl", message="Ingen modtagerliste valgt.")
            self.destroy()
            return
        
        if not os.path.exists(self.recipient_path):
            self.main_application.open_popup_window(size=[300, 125], title="Fejl", message="Den valgte modtagerliste eksisterer ikke.")
            self.destroy()
            return
        
        if not self.main_application.letter_type_frame.get() in self.main_application.letter_type_frame.brevtyper:
            self.main_application.open_popup_window(size=[300, 125], title="Fejl", message="Brevtype er ikke valgt eller ugyldig.")
            self.destroy()
            return

        # Position the window in the center of the root window
        x = self.root.winfo_x() + self.root.winfo_width() // 2 - self.size[0] // 2
        y = self.root.winfo_y() + self.root.winfo_height() // 2 - self.size[1] // 2
        self.geometry('+%d+%d' % (x, y))
        
        self.frame = ctk.CTkScrollableFrame(self, width=self.size[0], height=self.size[1])
        self.frame.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        
        self.file = pd.read_excel(self.recipient_path)
        self.file = self.file.dropna()
        
        if self.letter_type == '4-mdrs. brev' or self.letter_type == 'Gradbrev (stud.)':
            self.recipients = utils.get_students(self.recipient_path)
        elif self.letter_type == 'Gradbrev (vejl.)':
            self.recipients = utils.get_vejledere(self.recipient_path)
        else:
            self.main_application.open_popup_window(size=[300, 125], title="Fejl", message="Ugyldig brevtype.")
            self.destroy()
            return
        
        self.secretaries = utils.get_secretaries(self.letter_type)
        
        for institute in self.recipients[1]:
            if not institute in self.secretaries.keys():
                self.main_application.open_popup_window(size=[300, 125], title="Fejl", 
                                                        message=f"Kunne ikke finde sekretær for {institute} i modtagerliste.\n"
                                                        +"Dette kan skyldes at du har valgt forskellig brevtype og modtagerliste.\n"
                                                        +"Tjek om instituttet er stavet korrekt i modtagerliste, eller om du har valgt den rigtige modtagerliste.")
                self.destroy()
                return
        
        # Configure the window title
        self.num_recipients = len(self.recipients[0])
        self.title(f'Liste over {self.num_recipients} modtagere')

        
        self.to_label = ctk.CTkLabel(self.frame, text='Modtagere:', font=ctk.CTkFont(size=20, weight='bold'))
        self.to_label.grid(row=0, column=0, padx=(10,20), pady=10, sticky='nsw')
        self.cc_label = ctk.CTkLabel(self.frame, text='CC:', font=ctk.CTkFont(size=20, weight='bold'))
        self.cc_label.grid(row=0, column=1, padx=(20,10), pady=10, sticky='nsw')
        
        row = 1
        for i in range(self.num_recipients):
            to_label = ctk.CTkLabel(self.frame, text=self.recipients[0][i], font=ctk.CTkFont(size=12))
            to_label.grid(row=row, column=0, padx=(10,20), pady=5, sticky='nsw')
            cc_label = ctk.CTkLabel(self.frame, text=self.secretaries[self.recipients[1][i]], font=ctk.CTkFont(size=12))
            cc_label.grid(row=row, column=1, padx=(20,10), pady=5, sticky='nsw')
            row += 1
            
        self.frame.pack(side='top', fill='both', expand=True)


class PopupWindow(ctk.CTkToplevel):
    def __init__(self, root, size, title, message) -> None:
        super().__init__()

        self.root = root
        self.size = size
        self.window_title = title
        self.message = message
        
        # Configure the main window
        self.title(self.window_title)
        
        # Position the window in the center of the root window
        x = self.root.winfo_x() + self.root.winfo_width() // 2 - self.size[0] // 2
        y = self.root.winfo_y() + self.root.winfo_height() // 2 - self.size[1] // 2
        self.geometry('+%d+%d' % (x, y))
        
        
        # Configure the label
        self.label = ctk.CTkLabel(self, text=self.message, font=ctk.CTkFont(size=20, weight='bold'))
        self.label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        
        # Configure the button
        self.button = ctk.CTkButton(self, text='OK', width=50, height=25, command=self.destroy)
        self.button.grid(row=1, column=0, padx=10, pady=10)
        
        self.wm_transient(self.root)
    


class SendPopupWindow(ctk.CTkToplevel):
    def __init__(self, root, size, unresolved_names) -> None:
        super().__init__()

        self.root = root
        self.size = size
        self.unresolved_names = unresolved_names
        
        # Configure the main window
        self.title('Advarsel: Kunne ikke finde mail')
        
        # Position the window in the center of the root window
        x = self.root.winfo_x() + self.root.winfo_width() // 2 - self.size[0] // 2
        y = self.root.winfo_y() + self.root.winfo_height() // 2 - self.size[1] // 2
        self.geometry('+%d+%d' % (x, y))
        
        
        # Configure the label
        self.label = ctk.CTkLabel(self, text=f'Kunne ikke finde mail på {len(unresolved_names)} modtagere:', font=ctk.CTkFont(size=20, weight='bold'))
        self.label.grid(row=0, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)
        
        # Configure the scrollable frame
        self.frame = ctk.CTkScrollableFrame(self, width=self.size[0]-20)
        self.frame.grid(row=1, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

        for i, unresolved_name in enumerate(unresolved_names):
            label = ctk.CTkLabel(self.frame, text=unresolved_name, font=ctk.CTkFont(size=12))
            label.grid(row=i, column=0, padx=5, pady=5, sticky='nswe')


        # Configure the buttons
        self.save_unresolved_button = ctk.CTkButton(self, text='Gem uafklarede modtagere', width=250, height=50, command=self.save_unresolved)
        self.save_unresolved_button.grid(row=2, column=0, padx=10, pady=10, columnspan=2)

        self.continue_button = ctk.CTkButton(self, text='Fortsæt med at sende alligevel', width=125, height=50, command=self.send_anyway)
        self.continue_button.grid(row=3, column=0, padx=10, pady=10)

        self.continue_button = ctk.CTkButton(self, text='Annuler afsending', width=125, height=50, command=self.cancel)
        self.continue_button.grid(row=3, column=1, padx=10, pady=10)
        
        self.wm_transient(self.root)
    

    def save_unresolved(self):
        file = ctk.filedialog.asksaveasfile(
            mode='w',
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            parent=self
        )
        if file is None:
            return
        
        df = pd.DataFrame({'Modtagere': self.unresolved_names,
                           'Sekretærer': self.root.unresolved_secretaries,
                           'Vedhæftede filer': self.root.unresolved_vedh_filer})
        if self.root.letter_type == 'Gradbrev (vejl.)':
            df['Students'] = self.root.unresolved_students

        df.to_excel(file.name, index=False)
        file.close()

    def send_anyway(self):
        self.root.send_breve()
        self.destroy()
    
    def cancel(self):
        self.destroy()


if __name__ == '__main__':
    ctk.set_appearance_mode('Dark') # 'System', 'Light', 'Dark'
    ctk.set_default_color_theme('blue') # 'blue', 'green', 'dark-blue'
    app = MainApplication()
    app.mainloop()

