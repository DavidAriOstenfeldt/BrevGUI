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
        self.geometry(f'{600}x{625}')

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
        
        self.update_send_button()
     
        
    def set_content(self, _current_value=None):
        brev_type = self.letter_type_frame.get()

        if brev_type not in self.letter_type_frame.brevtyper:
            self.open_popup_window(size=[300, 125], title="Fejl", message="Brevtype er ikke valgt eller ugyldig.")
            return
        
        if self.language_frame.get() == 'auto':
            # TODO: implement language detection (not sure if it has to be here)
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
            

    def update_send_button(self):
        # Do a manual check, to not incur errors
        if self.letter_type_frame.brevtype_var.get() == "Vælg brevtype" or self.file_frame.file_button._text == 'Vælg mappe med breve' or self.recipient_frame.recipient_button._text == 'Vælg modtager liste':
            self.send_button_frame.send_button.configure(state='disabled')
        else:
            self.send_button_frame.send_button.configure(state='normal')

    
    def send_breve(self, type='auto'):
        if self.sender_frame.get() is None:
            self.open_popup_window(size=[300, 125], title="Fejl", message="Afsender navn ikke indtastet.")
            return
            
        
        letter_type = self.letter_type_frame.get()
        if letter_type == '4-mdrs. brev' or letter_type == 'Gradbrev (stud.)':
            recipients = utils.get_students(self.recipient_frame.get())
        elif letter_type == 'Gradbrev (vejl.)':
            recipients = utils.get_vejledere(self.recipient_frame.get())
            students = utils.get_students(self.recipient_frame.get())
            
        secretaries = utils.get_secretaries(self.letter_type_frame.get())
        
        vedh_filer_path = self.file_frame.get()
        vedh_filer = utils.get_vedh_filer(vedh_filer_path)
        
        num_recipients = len(recipients[0])
        for i in range(num_recipients):
            subject_text = self.subject_frame.get()
            content_text = self.content_frame.get()
            
            # Replace placeholder text
            if letter_type == 'Gradbrev (vejl.)':
                subject_text = subject_text.replace('[student navn]', students[0][i])
                content_text = content_text.replace('[student navn]', students[0][i])
            
            content_text = content_text.replace('[modtager navn]', recipients[0][i])
            content_text = content_text.replace('[afsender navn]', self.sender_frame.get())
            
            try:
                cc = secretaries[recipients[1][i]]
            except:
                self.open_popup_window(size=[300, 125], title="Fejl", message=f"Kunne ikke finde sekretær for {recipients[1][i]} i modtagerliste.\nTjek om du har valgt den rigtige brevtype og modtagerliste.")
                return

            vedh_fil = vedh_filer[i]
            
            if type == 'auto':
                unresolved = utils.send_email(recipients[0][i], cc, subject_text, content_text, vedh_fil)

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
        self.file_button.configure(text=value)
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
        self.recipient_button.configure(text=value)
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
        # TODO: Implement
        self.master.send_breve(type=self.get_send_mode())


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
        

if __name__ == '__main__':
    ctk.set_appearance_mode('Dark') # 'System', 'Light', 'Dark'
    ctk.set_default_color_theme('blue') # 'blue', 'green', 'dark-blue'
    app = MainApplication()
    app.mainloop()

