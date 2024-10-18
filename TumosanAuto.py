import os
import json
import subprocess
import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox
from PIL import ImageTk, Image

class ProgramInstaller:
    def __init__(self, window):
        self.window = window
        self.window.title('Tumosan USB Auto Program Yükleyici')

        self.usb_drive_path = self.find_usb_drive()

        usb_drive = self.find_usb_drive()
        self.bg_frame = Image.open(os.path.join(usb_drive, 'Kurulum', 'images', 'TumosanSol.png'))
        photo = ImageTk.PhotoImage(self.bg_frame)
        self.bg_panel = Label(self.window, image=photo, anchor='w')
        self.bg_panel.image = photo
        self.bg_panel.pack(fill='both', expand='yes')

        self.lgn_frame = Frame(self.window, bg='#040405', width=850, height=1070)
        self.lgn_frame.place(x=1150, y=0)

        self.txt = "USB Auto Program Yükleyici"
        self.heading = Label(self.lgn_frame, text=self.txt, font=('yu gothic ui', 17, "bold"), bg="#040405",
                             fg='white', bd=5, relief=FLAT)
        self.heading.place(x=80, y=30, width=300, height=30)

        self.program_listesi = Listbox(self.lgn_frame, selectmode=MULTIPLE, font=('yu gothic ui', 12), bg="#040405", fg="white")
        self.program_listesi.place(x=80, y=100, width=500, height=850)

        self.program_sil_dugme = tk.Button(self.lgn_frame, text="Program Sil", font=("yu gothic ui", 13, "bold"), width=15,
                                           bd=0, bg='#ff5733', cursor='hand2', activebackground='#ff5733', fg='white', command=self.program_sil)
        self.program_sil_dugme.place(x=383, y=960)

        self.yukle_programlar()

        self.program_yukle_dugme = tk.Button(self.lgn_frame, text="Yükle", font=("yu gothic ui", 13, "bold"), width=10,
                                             bd=0, bg='#3047ff', cursor='hand2', activebackground='#3047ff', fg='white', command=self.program_yukle)
        self.program_yukle_dugme.place(x=78, y=960)

        self.program_ekle_dugme = tk.Button(self.lgn_frame, text="Program Ekle", font=("yu gothic ui", 13, "bold"), width=15,
                                            bd=0, bg='green', cursor='hand2', activebackground='#ff5733', fg='white', command=self.program_ekle)
        self.program_ekle_dugme.place(x=218, y=960)
        
        self.programlar_dict = {}

        self.load_program_data()

    def find_usb_drive(self):
        drives = ['%s:' % d for d in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ']
        for drive in drives:
            if os.path.exists(os.path.join(drive, 'Kurulum')):
                return os.path.abspath(drive)  # USB sürücüsünün tam yolu
        return None

    def program_ekle(self):
        dosya_yolu = filedialog.askopenfilename(title="Program Seç", filetypes=(("Executable Files", "*.exe"), ("All Files", "*.*")))
        
        if dosya_yolu:
            program_adi = os.path.basename(dosya_yolu).split(".exe")[0]
            self.program_listesi.insert(tk.END, program_adi)
            self.programlar_dict[program_adi] = os.path.abspath(dosya_yolu)
            
            if self.usb_drive_path:
                self.save_program_data()

    def yukle_programlar(self):
        if self.usb_drive_path:
            json_dosya_yolu = os.path.join(self.usb_drive_path, 'Kurulum', 'programs.json')
            if os.path.exists(json_dosya_yolu):
                with open(json_dosya_yolu, 'r') as json_file:
                    self.programlar_dict = json.load(json_file)
                    self.program_listesi.delete(0, tk.END)  
                    for program in self.programlar_dict.keys():
                        self.program_listesi.insert(tk.END, program)

    def program_sil(self):
        secili_programlar = self.program_listesi.curselection()
        if not secili_programlar:
            messagebox.showerror("Hata", "Lütfen silmek istediğiniz programı seçin.")
            return

        secili_program = secili_programlar[0]
        secili_program_adi = self.program_listesi.get(secili_program)

        if secili_program_adi in self.programlar_dict:  
            self.program_listesi.delete(secili_program)
            del self.programlar_dict[secili_program_adi]  

            usb_drive = self.find_usb_drive()
            if usb_drive:
                self.save_program_data()  

            
            self.program_listesi.delete(0, tk.END)  
            for program in self.programlar_dict.keys():
                self.program_listesi.insert(tk.END, program)
        else:
            messagebox.showerror("Hata", f"Seçili program '{secili_program_adi}' sözlükte bulunamadı.")

    def update_program_paths(self, usb_drive):
        usb_drive_letter = usb_drive[0]

        for program_adi, program_yolu in self.programlar_dict.items():
            for drive_letter in range(ord('D'), ord('G') + 1):
                if program_yolu.startswith(chr(drive_letter) + ':'):
                    new_program_yolu = program_yolu.replace(chr(drive_letter) + ':', usb_drive_letter + ':')
                    self.programlar_dict[program_adi] = new_program_yolu

        self.save_program_data()

    def program_yukle(self):
        usb_drive = self.find_usb_drive() 

        if not usb_drive:
            messagebox.showerror("Hata", "USB sürücüsü bulunamadı.")
            return

        self.update_program_paths(usb_drive) 

        secilen_programlar = self.program_listesi.curselection()
        if not secilen_programlar:
            messagebox.showerror("Hata", "Lütfen en az bir program seçin.")
            return

        for program_indeks in secilen_programlar:
            program_adi = self.program_listesi.get(program_indeks)
            program_yolu = self.programlar_dict.get(program_adi)
            if not program_yolu:
                messagebox.showerror("Hata", f"{program_adi} programının yolu belirtilmemiş.")
                continue

            print(f"{program_adi} yükleniyor... Yol: {program_yolu}")

            try:
                process = subprocess.Popen([program_yolu], shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NEW_CONSOLE)
                process.wait()

                stderr_output = process.stderr.read().decode('utf-8')
                if process.returncode != 0:
                    if "Access is denied" in stderr_output:
                        devam_edecek_mi = messagebox.askyesno("Uyarı", f"Bu programı yüklemek için yeterli izniniz yok: {program_adi}. Devam etmek istiyor musunuz?")
                    else:
                        devam_edecek_mi = messagebox.askyesno("Hata", f"'{program_adi}' programı yüklenirken bir hata oluştu:\n{stderr_output}\n\nDevam etmek istiyor musunuz?")
                    
                    if not devam_edecek_mi:
                        break
                else:
                    print(f"{program_adi} yüklendi.")

            except Exception as e:
                messagebox.showerror("Hata", f"{program_adi} yüklenirken bir hata oluştu: {e}")

        messagebox.showinfo("Bilgi", "Seçilen programlar başarıyla yüklendi.")


    def save_program_data(self):
        if self.usb_drive_path:
            json_dosya_yolu = os.path.join(self.usb_drive_path, 'Kurulum', 'programs.json')
            with open(json_dosya_yolu, 'w') as json_file:
                json.dump(self.programlar_dict, json_file, indent=4)

    def load_program_data(self):
        if self.usb_drive_path:
            json_dosya_yolu = os.path.join(self.usb_drive_path, 'Kurulum', 'programs.json')
            if os.path.exists(json_dosya_yolu):
                with open(json_dosya_yolu, 'r') as json_file:
                    self.programlar_dict = json.load(json_file)
                    self.program_listesi.delete(0, tk.END)
                    for program in self.programlar_dict.keys():
                        self.program_listesi.insert(tk.END, program)

def sayfa():
    pencere = Tk()
    app = ProgramInstaller(pencere)
    pencere.geometry(f"{pencere.winfo_screenwidth()}x{pencere.winfo_screenheight()}")

    usb_drive = app.find_usb_drive()
    if usb_drive:
        app.update_program_paths(usb_drive)  
    else:
        messagebox.showerror("Hata", "USB sürücüsü bulunamadı.")

    pencere.mainloop()

if __name__ == '__main__':
    sayfa()