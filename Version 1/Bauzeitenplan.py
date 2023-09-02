from tkinter import *
from tkinter import ttk
import os

#GUI dev
PATH = os.path.dirname(os.path.realpath(__file__))
__version__ = '0.1'
__author__ = 'Justin'
hg = '#414245'
vg = 'white'
monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"]
font1 = "(\"Helvetica\", 20)"


class GUI(Tk):
    def __init__(self, geo):
        super().__init__()
        self.ui()
        self.geometry(geo)
        self.title('Bauzeitenplan Creator')
        self.resizable(0, 0)
        self.iconbitmap(PATH + '/icon.ico')

    def ui(self):
        self.configure(bg=hg)
        version = Label(master=self, text="Version  " + __version__, font=20, fg=vg, bg=hg)
        version.place(x=1793, y=1010, width=120, height=20)
        entwickler = Label(master=self, text=__author__, font=20, fg=vg, bg=hg)
        entwickler.place(x=1800, y=1030, width=100, height=20)
        logopfad = PhotoImage(PATH + "/logo.gif")
        logo = Label(master=self, image=logopfad, compound="center")
        logo.place(x=400, y=50, width=240, height=140)
        logo.image = logopfad

        # Eingabe
        Jahr_anfang_txt = Label(master=self, text="Anfangs Jahr", font=20, fg=vg, bg=hg)
        Jahr_anfang_txt.place(x=300, y=360, width=200, height=35)
        Jahr_anfang_eingabe = Entry(master=self, font=20, fg=vg, bg=hg)
        Jahr_anfang_eingabe.place(x=300, y=400, width=200, height=35)

        Jahr_ende_txt = Label(master=self, text="End Jahr", font=20, fg=vg, bg=hg)
        Jahr_ende_txt.place(x=520, y=360, width=200, height=35)
        Jahr_ende_eingabe = Entry(master=self, font=20, fg=vg, bg=hg)
        Jahr_ende_eingabe.place(x=520, y=400, width=200, height=35)

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TCombobox", fieldbackground=hg, background=hg)

        Monat_anfang_txt = Label(master=self, text="Anfangs Monat", font=20, fg=vg, bg=hg)
        Monat_anfang_txt.place(x=300, y=460, width=200, height=35)
        Monat_anfang_eingabe = ttk.Combobox(self, values=monate, font=20)
        Monat_anfang_eingabe.place(x=300, y=500, width=200, height=35)
        Monat_anfang_eingabe['state'] = 'readonly'

        Monat_ende_txt = Label(master=self, text="End Monat", font=20, fg=vg, bg=hg)
        Monat_ende_txt.place(x=520, y=460, width=200, height=35)
        Monat_ende_eingabe = ttk.Combobox(self, values=monate, font=20)
        Monat_ende_eingabe.place(x=520, y=500, width=200, height=35)
        Monat_ende_eingabe['state'] = 'readonly'

        Bauherr_txt = Label(master=self, text="Bauherr", font=20, fg=vg, bg=hg)
        Bauherr_txt.place(x=400, y=600, width=200, height=35)
        Bauherr_eingabe = Entry(master=self, font=20, fg=vg, bg=hg)
        Bauherr_eingabe.place(x=300, y=640, width=420, height=35)

        pname_txt = Label(master=self, text="Projektname", font=20, fg=vg, bg=hg)
        pname_txt.place(x=400, y=700, width=200, height=35)
        pname_eingabe = Entry(master=self, font=20, fg=vg, bg=hg)
        pname_eingabe.place(x=300, y=740, width=420, height=35)

        pnummer_txt = Label(master=self, text="Projektnummer", font=20, fg=vg, bg=hg)
        pnummer_txt.place(x=300, y=800, width=200, height=35)
        pnummer_eingabe = Entry(master=self, font=20, fg=vg, bg=hg)
        pnummer_eingabe.place(x=300, y=840, width=200, height=35)

        checkbox = IntVar()
        button_checkbox = Checkbutton(master=self, text='Baukosten', variable=checkbox, onvalue=1, offvalue=0, font=20, bg=hg)
        button_checkbox.place(x=800, y=490, width=120, height=40)

        oeffne_Datei = Button(master=self, text="Datei anzeigen", font=20, fg=hg, bg=vg)
        oeffne_Datei.place(x=1000, y=840, width=200, height=40)

if __name__ == "__main__":
    gui = GUI("1920x1080")
    gui.mainloop()

