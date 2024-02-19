from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import numpy as np
import os
from scrollable_frame import ScrollFrame
from send_breve import send_email, get_secretaries, get_students, get_vedh_filer, get_vejledere, display_email, get_final_recommendations, levenstein_distance

def browse_files(
    text_to_be_changed, type="single", title="Vælg fil", fallback_text="Vælg fil"
):
    if type == "single":
        filename = filedialog.askopenfilename(
            initialdir="/",
            title=title,
            filetypes=(
                ("all files", "*.*"),
                ("Excel files", "*.xlsx"),
                ("Text files", "*.txt*"),
            ),
        )
    elif type == "directory":
        filename = filedialog.askdirectory(initialdir="/", title=title)
    else:
        filename = ""

    if not filename:
        text_to_be_changed.configure(text=fallback_text)
    else:
        # filename_no_path = filename.split("/")[len(filename.split("/")) - 1]
        text_to_be_changed.configure(text=filename)


def set_text_of_email(text, entry, email="4-mdrs. brev", sprog="en"):
    text.delete(1.0, "end")
    entry.delete(0, "end")
    content_folder = "content/"
    if email == "4-mdrs. brev":
        if sprog == "dk":
            tekst_fil = open(content_folder + "tekst-da-kommende-afslutning.txt", "r", encoding='utf-8')
            entry.insert("end", "Kommende afslutning af ph.d.-studie")
        elif sprog == "en":
            tekst_fil = open(content_folder + "tekst-en-kommende-afslutning.txt", "r")
            entry.insert("end", "Upcoming completion of PhD study")
        text.insert("end", tekst_fil.read())

    elif email == "Gradbreve (Stud.)":
        if sprog == "dk":
            tekst_fil = open(content_folder + "tekst-da-grad-stud.txt", "r", encoding='utf-8')
            entry.insert("end", "Tildeling af grad")
        elif sprog == "en":
            tekst_fil = open(content_folder + "tekst-en-grad-stud.txt", "r", encoding='utf-8')
            entry.insert("end", "Award of the PhD Degree")
        text.insert("end", tekst_fil.read())
    elif email == "Gradbreve (Vejl.)":
        if sprog == "dk":
            tekst_fil = open(content_folder + "tekst-da-grad-vejl.txt", "r", encoding='utf-8')
            entry.insert("end", "Tildeling af grad til [student navn]")
        elif sprog == "en":
            tekst_fil = open(content_folder + "tekst-en-grad-vejl.txt", "r", encoding='utf-8')
            entry.insert("end", "Award of the PhD Degree to [student navn]")
        text.insert("end", tekst_fil.read())
    else:
        text.insert("end", "Something bad happened ....")


def open_new_window(root, modtager_liste, brev_type="4-mdrs. brev"):
    if modtager_liste == "" or modtager_liste == "Vælg modtager liste" or modtager_liste == "Vælg fil" or modtager_liste == "Vælg vedhæftede filer":
        return

    window = Toplevel(root)

    window.geometry("400x400")
    frame = ScrollFrame(window)

    file = pd.read_excel(modtager_liste)
    file = file.dropna()
    students = get_students(modtager_liste)
    secretaries = get_secretaries(brev_type=brev_type)
    window.title("Liste over %s modtagere" % len(students[0]))
    row = 0
    for i in range(len(students[0])):
        Label(frame.viewPort, text=f'to: {students[0][i]} cc: {secretaries[students[1][i]]}', padx=10,pady=5).grid(row=row, column=0, sticky="w")
        row += 1

    frame.pack(side="top", fill="both", expand=True)


def open_error_window(root, error_text, window_title='Fejl', size="200x30"):
    window = Toplevel(root)
    window.geometry(size)
    window.title(window_title)
    Label(window, text=error_text).pack()


def update_send_button_color(send_breve, afsender_navn, vedh_filer, modt_liste):
    if not afsender_navn.get() or vedh_filer['text'] == "Vælg vedhæftede filer" or modt_liste['text'] == "Vælg modtager liste":
        send_breve.config(bg='#f69697', activebackground='#f69697')
    else:
        send_breve.config(bg='#77DD77', activebackground='#77DD77')


def send_breve(root, send_button, modtager_liste, vedh_filer, afsender_navn, brev_type, emne, indhold):
    if send_button['bg'] == '#f69697':
        open_error_window(root, "Du mangler at udfylde nogle felter")
        return
    
    students = get_students(modtager_liste)
    secretaries = get_secretaries(brev_type=brev_type) #secretaries = get_secretaries(secretary_path, type=brev_type)
    
    if vedh_filer != "Vælg vedhæftede filer":
        vedh_filer_paths = get_vedh_filer(vedh_filer)
        if len(vedh_filer_paths) != len(students[0]):
            open_error_window(root, "Antallet af vedhæftede filer passer ikke med antallet af modtagere", size="400x30")
            return
        if brev_type == "Gradbreve (Stud.)":
            final_rec_dir = os.path.dirname(vedh_filer)
            final_rec_dir = os.path.join(final_rec_dir, "final recommendations")
            final_recommendations = get_final_recommendations(final_rec_dir)
            if len(final_recommendations) != len(students[0]):
                open_error_window(root, "Antallet af final recommendation filer passer ikke med antallet af modtagere", size="400x30")
                return

    if brev_type == "Gradbreve (Vejl.)":
        vejledere = get_vejledere(modtager_liste)
    
    unresolved = []
    for i in range(len(students[0])):
        emne_text = emne
        indhold_text = indhold    
        # Replace all placeholders in email with correct names
        emne_text = emne_text.replace('[student navn]', students[0][i])
        indhold_text = indhold_text.replace('[student navn]', students[0][i])
        if brev_type == "Gradbreve (Vejl.)":
            indhold_text = indhold_text.replace('[modtager navn]', vejledere[0][i])
            modtager = vejledere[0][i]
        else:
            indhold_text = indhold_text.replace('[modtager navn]', students[0][i])
            modtager = students[0][i]
        
        indhold_text = indhold_text.replace('[afsender navn]', afsender_navn)

        try:
            cc = secretaries[students[1][i]]
        except KeyError:
            open_error_window(root, f'Kunne ikke finde sekretær for {students[0][i]}\nTjek om du har valgt den rigtige brevtype og modtagerliste', size="325x40")
            return
        
        cc = secretaries[students[1][i]]
        vedh_fil = vedh_filer_paths[i]
        

        if brev_type == "Gradbreve (Stud.)":
            best_match_score = 0
            for recommendation in final_recommendations:
                score = 0
                for name in students[0][i].split(' '):
                    if name in recommendation:
                        score += 1
                        
                if score > best_match_score:
                    best_match_score = score
                    best_match = recommendation

            vedh_fil = vedh_filer_paths[i] + ';' + best_match

        unresolved_ = display_email(modtager, cc, emne_text, indhold_text, vedh_fil) # For testing purposes
        # unresolved_ = send_email(modtager, cc, emne_text, indhold_text, vedh_fil)
        unresolved += unresolved_

    if len(unresolved) > 0:
        open_error_window(root, f'Kunne ikke sende mail til {len(unresolved)} modtagere:\n{unresolved}', size="400x40")
        
    open_error_window(root, f'Sendte mail til {len(students[0]) - len(unresolved)} modtagere', window_title='Succes', size="325x40")


def BrevGUI():
    root = Tk()
    root.geometry("600x625")
    root.title("BrevGUI")

    frm = ttk.Frame(root)
    frm.grid()

    # Vælg brevtype
    brev_label = Label(frm, text="Brevtype")
    brev_options = ["4-mdrs. brev", "Gradbreve (Stud.)", "Gradbreve (Vejl.)"]
    clicked = StringVar()
    clicked.set(brev_options[0])

    brev_type = OptionMenu(
        frm,
        clicked,
        *brev_options,
        command=lambda x: set_text_of_email(email_text, emne_entry, email=clicked.get(), sprog=sprog.get())
    )
    brev_type.config(width=40, height=2)

    # Indtast navn
    afsender_label = Label(frm, text="Afsender navn")
    afsender_navn_var = StringVar()
    afsender_navn = Entry(frm, textvariable=afsender_navn_var, width=45)
    afsender_navn_var.trace('w', lambda *args: update_send_button_color(send_button, afsender_navn, vedh_filer, modt_liste))

    # Sprog
    sprog = StringVar()
    sprog_label = Label(frm, text="Sprog:")
    sprog_radio_en = Radiobutton(frm, text="Engelsk", variable=sprog, value="en", command=lambda: set_text_of_email(email_text, emne_entry, email=clicked.get(), sprog=sprog.get()))
    sprog_radio_dk = Radiobutton(frm, text="Dansk", variable=sprog, value="dk", command=lambda: set_text_of_email(email_text, emne_entry, email=clicked.get(), sprog=sprog.get()))
    sprog_radio_en.select()

    # Vedhæftede filer
    path_vedh_filer = ""
    vedh_filer_label = Label(frm, text="Vedhæftede filer")
    vedh_filer = Button(
        frm,
        text="Vælg vedhæftede filer",
        command=lambda: [browse_files(
            vedh_filer,
            "directory",
            "Vælg mappe med vedhæftede filer",
            "Vælg vedhæftede filer",
        ), update_send_button_color(
            send_button, afsender_navn, vedh_filer, modt_liste
        )],
        width=39,
        height=2,
    )

    # Modtager liste
    path_modtager_liste = ""
    modt_liste_label = Label(frm, text="Modtager liste")
    modt_liste = Button(
        frm,
        text="Vælg modtager liste",
        command=lambda: [browse_files(
            modt_liste, "single", "Vælg excel-ark med modtagere", "Vælg modtager liste"
        ), update_send_button_color(
            send_button, afsender_navn, vedh_filer, modt_liste
        )],
        width=39,
        height=2
    )

    # Emne
    emne_label = Label(frm, text="Emne:")
    emne_entry = Entry(frm, width=85)

    # Tekst
    email_label = Label(frm, text="Tekst:")
    email_text = Text(frm, width=71, height=19)
    set_text_of_email(email_text, emne_entry, clicked.get())

    # Send
    send_button = Button(
        frm,
        text="SEND BREVE",
        width=30,
        height=2,
        bg='#f69697',
        activebackground='#f69697',
        command=lambda: send_breve(
            root, 
            send_button,
            modt_liste.cget('text'), 
            vedh_filer.cget('text'),
            afsender_navn_var.get(),
            clicked.get(),
            emne_entry.get(),
            email_text.get("1.0", "end")
        )
    )

    # Se modtagere
    se_modtagere_button = Button(frm, text="Se modtagere", width=30, height=2, command=lambda: open_new_window(root, modt_liste.cget("text"), clicked.get()))


    # Grid layout
    brev_label.grid(column=0, row=0, padx=10, pady=(10, 0), columnspan=2)
    brev_type.grid(column=0, row=1, padx=(10, 10), pady=(0, 10), columnspan=2)
    afsender_label.grid(column=2, row=0, pady=(5, 0), sticky="", columnspan=2)
    afsender_navn.grid(column=2, row=1, padx=(0, 10), sticky="n", columnspan=2)
    vedh_filer_label.grid(column=0, row=2, padx=10, columnspan=2)
    modt_liste_label.grid(column=2, row=2, padx=(0, 10), columnspan=2)
    vedh_filer.grid(column=0, row=3, padx=10, columnspan=2)
    modt_liste.grid(column=2, row=3, padx=(0, 10), columnspan=2)
    sprog_label.grid(column=0, row=4, padx=10, pady=10, sticky="w")
    sprog_radio_dk.grid(column=1, row=4, padx=0)
    sprog_radio_en.grid(column=2, row=4, padx=10)
    emne_label.grid(column=0, row=5, padx=10, pady=10, sticky="w")
    emne_entry.grid(column=1, row=5, padx=(0, 10), pady=10, sticky="e", columnspan=3)
    email_label.grid(column=0, row=6, padx=10, sticky="w", columnspan=4)
    email_text.grid(column=0, row=7, padx=10, columnspan=4)
    se_modtagere_button.grid(column=0, row=8, padx=10, pady=10, sticky="w", columnspan=2)
    send_button.grid(column=2, row=8, padx=10, pady=10, sticky="e", columnspan=2)

    root.mainloop()


# Press the green button in the gutter to run the script.
if __name__ == "__main__":
    BrevGUI()
