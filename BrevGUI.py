from tkinter import *
from tkinter import ttk
from tkinter import filedialog


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
        filename_no_path = filename.split("/")[len(filename.split("/")) - 1]
        text_to_be_changed.configure(text=filename_no_path)


def set_text_of_email(text, entry, email="4-mdrs. brev"):
    text.delete(1.0, "end")
    entry.delete(0, "end")
    if email == "4-mdrs. brev":
        text.insert("end", "EYOOO")
        entry.insert("end", "Kommende afslutning af ph.d.-studie")
    elif email == "Gradbreve (Stud.)":
        text.insert("end", "this is a gradbrev")
        entry.insert("end", "Tildeling af grad")
    elif email == "Gradbreve (Vejl.)":
        text.insert("end", "this is a gradbrev to a vejleder")
        entry.insert("end", "Tildeling af grad til ")
    else:
        text.insert("end", "Something bad happened ....")


def BrevGUI():
    root = Tk()
    root.geometry("600x600")
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
        command=lambda x: set_text_of_email(email_text, emne_entry, email=clicked.get())
    )
    brev_type.config(width=40, height=2)

    # Indtast navn
    afsender_label = Label(frm, text="Afsender navn")
    afsender_navn = Entry(frm, width=45)

    # Vedhæftede filer
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
    modt_liste_label = Label(frm, text="Modtager liste")
    modt_liste = Button(
        frm,
        text="Vælg modtager liste",
        command=lambda: browse_files(
            modt_liste, "single", "Vælg excel-ark med modtagere", "Vælg modtager liste"
        ),
        width=39,
        height=2,
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

    # Grid layout
    brev_label.grid(column=0, row=0, padx=10, pady=(10, 0))
    brev_type.grid(column=0, row=1, padx=(10, 10), pady=(0, 10))
    afsender_label.grid(column=1, row=0, pady=(5, 0), sticky="")
    afsender_navn.grid(column=1, row=1, padx=(0, 10), sticky="n")
    vedh_filer_label.grid(column=0, row=2, padx=10)
    modt_liste_label.grid(column=1, row=2, padx=(0, 10))
    vedh_filer.grid(column=0, row=3, padx=10)
    modt_liste.grid(column=1, row=3, padx=(0, 10))
    emne_label.grid(column=0, row=4, padx=10, pady=10, sticky="w")
    emne_entry.grid(column=0, row=4, padx=(0, 10), pady=10, sticky="e", columnspan=2)
    email_label.grid(column=0, row=5, padx=10, sticky="w")
    email_text.grid(column=0, row=6, padx=10, columnspan=2)
    send_button.grid(column=1, row=7, padx=10, pady=10, sticky="e")

    root.mainloop()


# Press the green button in the gutter to run the script.
if __name__ == "__main__":
    BrevGUI()
