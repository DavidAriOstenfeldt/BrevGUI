import tkinter as tk
import customtkinter as ctk


ctk.set_appearance_mode('Dark') # 'System', 'Light', 'Dark'
ctk.set_default_color_theme('blue') # 'blue', 'green', 'dark-blue'


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
        ### Configure the letter type frame ###
        self.brevtype_frame = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.brevtype_frame.grid(row=0, column=0, sticky='nsew')
        self.brevtype_frame.grid_rowconfigure(0, weight=1)
        self.brevtype_frame.grid_rowconfigure(1, weight=2)
        
        # Label
        self.brevtype_label = ctk.CTkLabel(self.brevtype_frame, text='Brevtype', font=ctk.CTkFont(size=20, weight='bold'))
        self.brevtype_label.grid(row=0, column=0, padx=10, pady=(10,5), sticky='nsew')

        # OptionMenu
        self.brevtyper = ['4-mdrs. brev', 'Gradbrev (stud.)', 'Gradbrev (vejl.)']
        self.brevtype_var = tk.StringVar()
        self.brevtype_option_menu = ctk.CTkOptionMenu(self.brevtype_frame, variable=self.brevtype_var, command=self.set_content, values=self.brevtyper, width=275)
        self.brevtype_option_menu.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')
        self.brevtype_option_menu.set('Vælg brevtype')

        ### Configure the sender frame ###
        self.sender_frame = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.sender_frame.grid(row=0, column=1, sticky='nsew')
        self.sender_frame.grid_rowconfigure(0, weight=1)
        self.sender_frame.grid_rowconfigure(1, weight=2)

        # Label
        self.sender_label = ctk.CTkLabel(self.sender_frame, text='Afsender', font=ctk.CTkFont(size=20, weight='bold'))
        self.sender_label.grid(row=0, column=0, padx=10, pady=(10,5))

        # Entry
        self.sender_entry = ctk.CTkEntry(self.sender_frame, placeholder_text='Afsender', width=275)
        self.sender_entry.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')

        ##### ROW 2 #####
        ### Configure the file frame ###
        self.file_frame = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.file_frame.grid(row=1, column=0, sticky='nsew')
        self.file_frame.grid_rowconfigure(0, weight=1)
        self.file_frame.grid_rowconfigure(1, weight=2)

        # Label
        self.file_label = ctk.CTkLabel(self.file_frame, text='Vedhæftede filer', font=ctk.CTkFont(size=20, weight='bold'))
        self.file_label.grid(row=0, column=0, padx=10, pady=(10,5))

        # Button
        self.file_button = ctk.CTkButton(self.file_frame, text='Vælg mappe med breve', width=275)
        self.file_button.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')

        ### Configure the recipient frame ###
        self.recipient_frame = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.recipient_frame.grid(row=1, column=1, sticky='nsew')
        self.recipient_frame.grid_rowconfigure(0, weight=1)
        self.recipient_frame.grid_rowconfigure(1, weight=2)
        
        # Label
        self.recipient_label = ctk.CTkLabel(self.recipient_frame, text='Vælg modtager liste', font=ctk.CTkFont(size=20, weight='bold'))
        self.recipient_label.grid(row=0, column=0, padx=10, pady=(10,5))

        # Button
        self.recipient_button = ctk.CTkButton(self.recipient_frame, text='Vælg modtager liste', width=275)
        self.recipient_button.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')

        ##### ROW 3 #####
        ### Configure the language frame ###
        self.language_frame = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.language_frame.grid(row=2, column=0, columnspan=2, sticky='nsew')
        self.language_frame.grid_rowconfigure(0, weight=1)
        self.language_frame.grid_columnconfigure(0, weight=1)
        self.language_frame.grid_columnconfigure((1,2,3), weight=3)

        # Label
        self.language_label = ctk.CTkLabel(self.language_frame, text='Sprog:', font=ctk.CTkFont(size=20, weight='bold'))
        self.language_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsw')

        # Radiobuttons
        self.language_var = tk.StringVar()
        self.language_var.set('da')
        self.language_radiobutton1 = ctk.CTkRadioButton(self.language_frame, text='Dansk', command=self.set_content, variable=self.language_var, value='da')
        self.language_radiobutton1.grid(row=0, column=1, padx=(10,5), pady=10, sticky='nse')
        self.language_radiobutton2 = ctk.CTkRadioButton(self.language_frame, text='Engelsk', command=self.set_content, variable=self.language_var, value='en')
        self.language_radiobutton2.grid(row=0, column=2, padx=0, pady=10, sticky='nse')
        self.language_radiobutton3 = ctk.CTkRadioButton(self.language_frame, text='Automatisk', command=self.set_content, variable=self.language_var, value='auto')
        self.language_radiobutton3.grid(row=0, column=3, padx=(5,10), pady=10, sticky='nse')

        ##### ROW 4 #####
        ### Configure the subject frame ###
        self.subject_frame = ctk.CTkFrame(self, width=600, corner_radius=0)
        self.subject_frame.grid(row=3, column=0, columnspan=2, sticky='nsew')
        self.subject_frame.grid_rowconfigure(0, weight=1)
        self.subject_frame.grid_columnconfigure(0, weight=1)
        self.subject_frame.grid_columnconfigure(1, weight=3)

        # Label
        self.subject_label = ctk.CTkLabel(self.subject_frame, text='Emne:', font=ctk.CTkFont(size=20, weight='bold'))
        self.subject_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsw')

        # Entry
        self.subject_entry = ctk.CTkEntry(self.subject_frame, placeholder_text='Emne', width=500)
        self.subject_entry.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')

        ##### ROW 5 #####
        ### Configure the content frame ###
        self.content_frame = ctk.CTkFrame(self, width=600, corner_radius=0)
        self.content_frame.grid(row=4, column=0, columnspan=2, sticky='nsew')
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=3)
        self.content_frame.grid_columnconfigure(0, weight=1)

        # Label
        self.content_label = ctk.CTkLabel(self.content_frame, text='Indhold tekst:', font=ctk.CTkFont(size=20, weight='bold'))
        self.content_label.grid(row=0, column=0, padx=10, pady=10, sticky='nsw')

        # Text
        self.content_text = ctk.CTkTextbox(self.content_frame)
        self.content_text.grid(row=1, column=0, padx=10, pady=(0,10), sticky='nsew')

        ##### ROW 6 #####
        ### Configure the send button frame ###
        self.send_button_frame = ctk.CTkFrame(self, width=600, corner_radius=0)
        self.send_button_frame.grid(row=5, column=0, columnspan=2, sticky='nsew')
        self.send_button_frame.grid_rowconfigure(0, weight=1)
        self.send_button_frame.grid_columnconfigure((0,1,2), weight=1)

        # See recipients button
        self.see_recipients_button = ctk.CTkButton(self.send_button_frame, text='Se modtagere', width=250, height=50)
        self.see_recipients_button.grid(row=0, column=0, rowspan=2, padx=10, pady=10, sticky='nsw')

        # Send mode radio buttons
        self.send_mode_var = tk.StringVar()
        self.send_mode_var.set('Auto')
        self.send_mode_radiobutton1 = ctk.CTkRadioButton(self.send_button_frame, text='Auto', variable=self.send_mode_var, value='auto')
        self.send_mode_radiobutton1.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')
        self.send_mode_radiobutton2 = ctk.CTkRadioButton(self.send_button_frame, text='Manuel', variable=self.send_mode_var, value='manuel')
        self.send_mode_radiobutton2.grid(row=1, column=1, padx=10, pady=10, sticky='nsew')

        # Send Button
        self.send_button = ctk.CTkButton(self.send_button_frame, text='Send breve', width=250, height=50)
        self.send_button.grid(row=0, column=2, rowspan=2, padx=10, pady=10, sticky='nse')


    def set_content(self, _current_value=None):
        brev_type = self.brevtype_var.get()

        if brev_type not in self.brevtyper:
            print('Brev type er ikke valgt eller ugyldig')
            return
        
        if self.language_var.get() == 'auto':
            _language = 'da'
        else:
            _language = self.language_var.get()

        text_file = open('content/' + _language + '-' + brev_type + '.txt', 'r', encoding='utf-8')
        text = text_file.read()
        text_file.close()

        # Set subject
        subject = text.split('Emne: ')[1].split('\n')[0]
        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, subject)

        # Set content
        content = text.split('Indhold:\n')[1]
        self.content_text.delete(1.0, tk.END)
        self.content_text.insert(tk.END, content)
        





if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()

