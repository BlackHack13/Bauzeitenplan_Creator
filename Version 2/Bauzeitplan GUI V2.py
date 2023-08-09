import tkinter
import tkinter.messagebox
import customtkinter
from PIL import Image, ImageTk
import openpyxl
import os
from openpyxl import Workbook
from tkinter.filedialog import asksaveasfile


PATH = os.path.dirname(os.path.realpath(__file__))
customtkinter.set_appearance_mode("System")  # Modes: "System", "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue" "sweetkind"
monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"]
version = "Version: 2.0"

workbook = Workbook()
worksheet = workbook.active

class GUI(customtkinter.CTk):

    WIDTH = 800
    HEIGHT = 800

    def __init__(self):
        super().__init__()

        self.title("Bauzeitplan Creator")
        self.iconbitmap(PATH + '/icon.ico')
        self.geometry(f"{GUI.WIDTH}x{GUI.HEIGHT}")
        self.protocol(self.close_app)

        # ============ frames ============
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame_links = customtkinter.CTkFrame(master=self, width=400, corner_radius=0)
        self.frame_links.grid(row=0, column=0, sticky="nswe")

        self.frame_mitte = customtkinter.CTkFrame(master=self, width=180, corner_radius=0)
        self.frame_mitte.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        self.frame_rechts = customtkinter.CTkFrame(master=self, width=180, corner_radius=0)
        self.frame_rechts.grid(row=0, column=2, sticky="nswe")

        # ==== grid layout (1x11) L ===
        self.frame_links.grid_rowconfigure(0, minsize=10)   # empty row with minsize as spacing
        self.frame_links.grid_rowconfigure(5, weight=1)  # empty row as spacing
        self.frame_links.grid_rowconfigure(8, minsize=20)    # empty row with minsize as spacing
        self.frame_links.grid_rowconfigure(20, minsize=10)  # empty row with minsize as spacing

        # === grid layout (3x7) M ===
        self.frame_mitte.rowconfigure((0, 1, 2), weight=1)
        self.frame_mitte.rowconfigure(8, weight=10)
        self.frame_mitte.columnconfigure((0, 1), weight=1)
        self.frame_mitte.columnconfigure(2, weight=0)

        # === grid layout (1x11) R ===
        self.frame_rechts.rowconfigure(0, weight=1)
        self.frame_rechts.rowconfigure(7, weight=10)
        self.frame_rechts.columnconfigure(8, weight=1)
        self.frame_rechts.columnconfigure(2, weight=0)


        #////// Links
        logo = Image.open(PATH + "/Logo.gif")
        self.bg_logo = ImageTk.PhotoImage(logo)
        self.logo_label = tkinter.Label(master=self.frame_links, image=self.bg_logo)
        self.logo_label.grid(row=4, column=0)

        self.label_darstellung = customtkinter.CTkLabel(master=self.frame_links, text="Darstellung:")
        self.label_darstellung.grid(row=9, column=0)
        self.darstellung = customtkinter.CTkOptionMenu(master=self.frame_links, values=["Light", "Dark", "System"], command=self.change_appearance_mode)
        self.darstellung.grid(row=10, column=0, pady=10, padx=20)
        self.darstellung.set("System")

        # ////// Mitte
        self.bauherr_entry = customtkinter.CTkEntry(master=self.frame_mitte, placeholder_text="Bauherr: ", width=300, height=25, border_width=2, corner_radius=5)
        self.bauherr_entry.grid(row=1, column=0, columnspan=3)
        self.projektname_entry = customtkinter.CTkEntry(master=self.frame_mitte, placeholder_text="Projektname: ", width=300, height=25, border_width=2, corner_radius=5)
        self.projektname_entry.grid(row=2, column=0, columnspan=3)
        self.projektnummer_entry = customtkinter.CTkEntry(master=self.frame_mitte, placeholder_text="Projektnummer: ", width=300, height=25, border_width=2, corner_radius=5)
        self.projektnummer_entry.grid(row=3, column=0, columnspan=3)

        self.startjahr_entry = customtkinter.CTkEntry(master=self.frame_mitte, placeholder_text="Startjahr: ", width=120, height=25, border_width=2, corner_radius=5)
        self.startjahr_entry.grid(row=5, column=0, padx=0, pady=10)
        self.endjahr_entry = customtkinter.CTkEntry(master=self.frame_mitte, placeholder_text="Endjahr: ", width=120, height=25, border_width=2, corner_radius=5)
        self.endjahr_entry.grid(row=5, column=1, padx=0, pady=10)

        self.startmonat_entry = customtkinter.CTkOptionMenu(master=self.frame_mitte, values=monate)
        self.startmonat_entry.grid(row=6, column=0)
        self.endmonat_entry = customtkinter.CTkOptionMenu(master=self.frame_mitte, values=monate)
        self.endmonat_entry.grid(row=6, column=1)

        self.baukosten = customtkinter.CTkCheckBox(master=self.frame_mitte, text="Baukosten", onvalue="on", offvalue="off")
        self.baukosten.grid(row=7, column=0, pady=10)

        self.button_speichern = customtkinter.CTkButton(master=self.frame_mitte, text="Speichern", width=300, height=25, border_width=2, corner_radius=5, command=self.save_file)
        self.button_speichern.grid(row=9, column=0, pady=70)
        self.button_offnen = customtkinter.CTkButton(master=self.frame_mitte, text="Öffnen in Excel", width=300, height=25, border_width=2, corner_radius=5, command=self.open)
        self.button_offnen.grid(row=9, column=1, pady=70)

        # ////// Rechts
        self.preview = customtkinter.CTkLabel(master=self.frame_rechts, text="Preview")
        self.preview.grid(row=0, column=0, pady=0, padx=20)

        img_baukosten = Image.open(PATH + "/Logo.gif")
        self.bg_img_baukosten= ImageTk.PhotoImage(img_baukosten)
        self.img_baukosten = tkinter.Label(master=self.frame_rechts, image=self.bg_img_baukosten)
        self.img_baukosten.grid(row=2, column=0)

        img_o_baukosten = Image.open(PATH + "/Logo.gif")
        self.bg_img_o_baukosten = ImageTk.PhotoImage(img_o_baukosten)
        self.img_o_baukosten = tkinter.Label(master=self.frame_rechts, image=self.bg_img_o_baukosten)
        self.img_o_baukosten.grid(row=7, column=0)

        self.label_name = customtkinter.CTkLabel(master=self.frame_rechts, text="Justin").grid(row=8, column=0)
        self.label_version = customtkinter.CTkLabel(master=self.frame_rechts, text=version).grid(row=9, column=0, pady=10)

    def fehler_eingabe(self):   #TODO
        number1 = self.startjahr_entry.get()
        number2 = self.endjahr_entry.get()
        try:
            startjahr = int(number1)
            endjahr = int(number2)
            self.button_offnen.configure(state="normal")
            self.button_speichern.configure(state="normal")

        except ValueError:
            self.button_offnen.configure(state="disabled")
            self.button_speichern.configure(state="disabled")
            tkinter.messagebox.showerror("Fehler", "Ungültige Eingabe. Bitte geben Sie eine gültiges Jahr ein.")

    def save_file(self):
        projektnummer = self.projektnummer_entry.get()
        dateiname = projektnummer + " Bauzeitplan"
        filepath = tkinter.filedialog.asksaveasfilename(
            initialfile=dateiname,
            defaultextension=".xlsx",
            filetypes=[("Excel Arbeitsmappe", ".xlsx")])
        workbook.save(filepath)


    def open(self):
        os.startfile(PATH + "/Bauzeitplan.xlsx")

    def close_app(self):
        self.destroy()

    def change_appearance_mode(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)


if __name__ == "__main__":
    #Excel()
    gui = GUI()
    gui.mainloop()

