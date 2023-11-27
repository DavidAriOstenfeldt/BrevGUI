from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import numpy as np

from scrollable_frame import ScrollFrame

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
    if email == "4-mdrs. brev":
        if sprog == "dk":
            tekst_fil = open("tekst-da-kommende-afslutning.txt", "r", encoding='utf-8')
            entry.insert("end", "Kommende afslutning af ph.d.-studie")
        elif sprog == "en":
            tekst_fil = open("tekst-en-kommende-afslutning.txt", "r")
            entry.insert("end", "Upcoming completion of PhD study")
        text.insert("end", tekst_fil.read())

    elif email == "Gradbreve (Stud.)":
        if sprog == "dk":
            tekst_fil = open("tekst-da-grad-stud.txt", "r", encoding='utf-8')
            entry.insert("end", "Tildeling af grad")
        elif sprog == "en":
            tekst_fil = open("tekst-en-grad-stud.txt", "r", encoding='utf-8')
            entry.insert("end", "Award of the PhD Degree")
        text.insert("end", tekst_fil.read())
    elif email == "Gradbreve (Vejl.)":
        if sprog == "dk":
            tekst_fil = open("tekst-da-grad-vejl.txt", "r", encoding='utf-8')
            entry.insert("end", "Tildeling af grad til [student navn]")
        elif sprog == "en":
            tekst_fil = open("tekst-en-grad-vejl.txt", "r", encoding='utf-8')
            entry.insert("end", "Award of the PhD Degree to [student navn]")
        text.insert("end", tekst_fil.read())
    else:
        text.insert("end", "Something bad happened ....")


def open_new_window(root, modtager_liste):
    if modtager_liste == "" or modtager_liste == "Vælg modtager liste" or modtager_liste == "Vælg fil" or modtager_liste == "Vælg vedhæftede filer":
        return

    window = Toplevel(root)

    window.geometry("400x400")
    frame = ScrollFrame(window)

    file = pd.read_excel(modtager_liste)
    file = file.dropna()
    students = np.array(file["Navn"])
    window.title("Liste over %s modtagere" % len(students))
    row = 0
    for student in students:
        Label(frame.viewPort, text=student,padx=10,pady=5).grid(row=row, column=0, sticky="w")
        row += 1

    frame.pack(side="top", fill="both", expand=True)




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
    afsender_navn = Entry(frm, width=45)

    # Sprog
    sprog = StringVar()
    sprog_label = Label(frm, text="Sprog:")
    sprog_radio_en = Radiobutton(frm, text="Engelsk", variable=sprog, value="en", command=lambda: set_text_of_email(email_text, emne_entry, email=clicked.get(), sprog=sprog.get()))
    sprog_radio_dk = Radiobutton(frm, text="Dansk", variable=sprog, value="dk", command=lambda: set_text_of_email(email_text, emne_entry, email=clicked.get(), sprog=sprog.get()))
    sprog_radio_en.select()

    # Vedhæftede filer
    path_vedh_filer = ""
    vedh_filer_label = Label(frm, text="Vedhæftede filer (valgfri)")
    vedh_filer = Button(
        frm,
        text="Vælg vedhæftede filer",
        command=lambda: browse_files(
            vedh_filer,
            "directory",
            "Vælg mappe med vedhæftede filer",
            "Vælg vedhæftede filer",
        ),
        width=39,
        height=2,
    )

    # Modtager liste
    path_modtager_liste = ""
    modt_liste_label = Label(frm, text="Modtager liste")
    modt_liste = Button(
        frm,
        text="Vælg modtager liste",
        command=lambda: browse_files(
            modt_liste, "single", "Vælg excel-ark med modtagere", "Vælg modtager liste"
        ),
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
        bg="#77DD77",
        activebackground="#77DD77",
    )

    # Se modtagere
    se_modtagere_button = Button(frm, text="Se modtagere", width=30, height=2, command=lambda: open_new_window(root, modt_liste.cget("text")))


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
