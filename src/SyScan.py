import io, os, time, threading, tempfile, pathlib, uuid
import customtkinter as ctk
import win32com.client
import pythoncom
from tkinter import messagebox
from fpdf import FPDF
from PIL import Image

# --- Configurações de Caminhos ---
C_PATH = r"C:\SyScan_Backup"
USER_PICTURES = str(pathlib.Path.home() / "Pictures" / "SyScan_Digitalizacoes")

for p in [C_PATH, USER_PICTURES]:
    os.makedirs(p, exist_ok=True)

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

def get_wia_devices():
    pythoncom.CoInitialize()
    try:
        wia = win32com.client.Dispatch("WIA.DeviceManager")
        return {dev.Properties("Name").Value: dev.DeviceID for dev in wia.DeviceInfos}
    except Exception: return {}
    finally: pythoncom.CoUninitialize()

def scan_to_file(device_id, pasta_temp, indice):
    pythoncom.CoInitialize()
    try:
        wia = win32com.client.Dispatch("WIA.DeviceManager")
        info = next(d for d in wia.DeviceInfos if d.DeviceID == device_id)
        dev = info.Connect()
        item = dev.Items[0]
        
        for prop in ["Horizontal Resolution", "Vertical Resolution"]:
            try: item.Properties(prop).Value = 300
            except: pass
            
        image = item.Transfer("{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")
        caminho = os.path.join(pasta_temp, f"pg_{indice:03d}_{uuid.uuid4().hex[:4]}.png")
        Image.open(io.BytesIO(image.FileData.BinaryData)).save(caminho, "PNG")
        return caminho
    except Exception as e:
        raise Exception(f"O scanner não respondeu.\nErro: {e}")
    finally: pythoncom.CoUninitialize()

class ScannerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ADAR - SyScaner")
        self.geometry("950x650")
        
        self.devices_ids = {}
        self.paginas_scaneadas = []
        self.preview_images_refs = [] # Guardar referências das imagens aqui
        self.path_root = os.path.dirname(os.path.abspath(__file__))
        
        self._setup_ui()
        self._load_icons()
        self.after(500, self.carregar)

    def _load_icons(self):
        icon_path = os.path.join(self.path_root, "logo.ico")
        img_path = os.path.join(self.path_root, "adar.png")
        if os.path.exists(icon_path):
            try: self.after(200, lambda: self.iconbitmap(icon_path))
            except: pass
        if os.path.exists(img_path):
            try:
                img = ctk.CTkImage(Image.open(img_path), size=(200, 50))
                self.logo_label.configure(image=img, text="")
            except: pass

    def _setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Painel Esquerdo
        self.left_pannel = ctk.CTkFrame(self, width=320, corner_radius=0)
        self.left_pannel.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        self.logo_label = ctk.CTkLabel(self.left_pannel, text="ADAR SYSCANER", font=("Roboto", 22, "bold"))
        self.logo_label.pack(pady=20)
        
        self.combo = ctk.CTkComboBox(self.left_pannel, width=280, state="readonly")
        self.combo.set("Buscando scanners...")
        self.combo.pack(pady=10)

        ctk.CTkButton(self.left_pannel, text="Atualizar Lista", command=self.carregar, 
                      fg_color="#34495e", height=35).pack(pady=5)
        
        self.btn_iniciar = ctk.CTkButton(self.left_pannel, text="DIGITALIZAR PÁGINA", height=60, width=280, 
                                        font=("Roboto", 16, "bold"), command=self.fluxo_digitalizacao,
                                        fg_color="#27ae60", hover_color="#219150")
        self.btn_iniciar.pack(pady=25)

        self.btn_remove = ctk.CTkButton(self.left_pannel, text="EXCLUIR ÚLTIMA", height=60, width=280, 
                                        font=("Roboto", 16, "bold"), fg_color="#e67e22", 
                                        state="disabled", command=self.remover_ultima)
        self.btn_remove.pack(pady=5)

        self.btn_finalizar = ctk.CTkButton(self.left_pannel, text="FINALIZAR", height=60, width=280, 
                                          font=("Roboto", 16, "bold"),
                                          command=self.finalizar_pdf, fg_color="#2980b9", state="disabled")
        self.btn_finalizar.pack(pady=20)

        self.status = ctk.CTkLabel(self.left_pannel, text="Iniciando...", font=("Roboto", 12))
        self.status.pack(side="bottom", pady=15)

        # Painel Direito
        self.preview_pannel = ctk.CTkScrollableFrame(self, corner_radius=15, fg_color="#ebebeb")
        self.preview_pannel.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        self.msg_label = ctk.CTkLabel(self.preview_pannel, text="O preview aparecerá aqui")
        self.msg_label.pack(pady=20)

    def carregar(self):
        def busca():
            self.devices_ids = get_wia_devices()
            nomes = list(self.devices_ids.keys())
            if nomes:
                self.after(0, lambda: [self.combo.configure(values=nomes), self.combo.set(nomes[0]), 
                                       self.status.configure(text="Scanner Pronto")])
            else:
                self.after(0, lambda: self.status.configure(text="Nenhum scanner detectado"))
        threading.Thread(target=busca, daemon=True).start()

    def fluxo_digitalizacao(self):
        if self.combo.get() in ["", "Buscando scanners...", "Nenhum scanner detectado"]: 
            return messagebox.showwarning("Aviso", "Selecione um scanner primeiro!")
        self.btn_iniciar.configure(state="disabled")
        self.status.configure(text="Digitalizando...")
        dev_id = self.devices_ids[self.combo.get()]
        threading.Thread(target=self.executar_captura, args=(dev_id, len(self.paginas_scaneadas)+1), daemon=True).start()

    def executar_captura(self, dev_id, indice):
        try:
            caminho = scan_to_file(dev_id, tempfile.gettempdir(), indice)
            self.paginas_scaneadas.append(caminho)
            self.after(10, self.atualizar_preview_completo)
        except Exception as e:
            self.after(10, lambda m=str(e): [messagebox.showerror("Erro", m), self.reset_ui()])

    def atualizar_preview_completo(self):
        # 1. Limpar painel
        for widget in self.preview_pannel.winfo_children():
            widget.destroy()
        self.preview_images_refs = []

        if not self.paginas_scaneadas:
            ctk.CTkLabel(self.preview_pannel, text="O preview aparecerá aqui").pack(pady=20)
            self.btn_finalizar.configure(state="disabled")
            self.btn_remove.configure(state="disabled")
        else:
            # 2. Reconstruir lista de imagens
            for caminho in self.paginas_scaneadas:
                try:
                    img = Image.open(caminho)
                    w, h = img.size
                    ratio = 500 / w
                    ctk_img = ctk.CTkImage(light_image=img, size=(500, int(h * ratio)))
                    
                    self.preview_images_refs.append(ctk_img) # Mantém a referência viva
                    lbl = ctk.CTkLabel(self.preview_pannel, image=ctk_img, text="")
                    lbl.pack(pady=10)
                except: pass
            
            self.btn_finalizar.configure(state="normal")
            self.btn_remove.configure(state="normal")

        self.status.configure(text=f"Total: {len(self.paginas_scaneadas)} página(s)")
        self.btn_iniciar.configure(state="normal")

    def remover_ultima(self):
        if self.paginas_scaneadas:
            caminho = self.paginas_scaneadas.pop()
            # Limpa preview antes de deletar arquivo
            for widget in self.preview_pannel.winfo_children():
                widget.destroy()
            self.update_idletasks()
            
            try:
                if os.path.exists(caminho): os.remove(caminho)
            except: pass
            
            self.atualizar_preview_completo()

    def finalizar_pdf(self):
        self.status.configure(text="Criando PDF...")
        self.btn_finalizar.configure(state="disabled")
        threading.Thread(target=self.gerar_pdf_process, daemon=True).start()

    def gerar_pdf_process(self):
        try:
            pdf = FPDF()
            for p in self.paginas_scaneadas:
                with Image.open(p).convert("RGB") as img:
                    w, h = img.size
                    ratio = min(210/w, 297/h)
                    nw, nh = w*ratio, h*ratio
                    with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
                        img.save(tmp.name, "JPEG", quality=80)
                        pdf.add_page()
                        pdf.image(tmp.name, (210-nw)/2, (297-nh)/2, nw, nh)
                    os.remove(tmp.name)

            nome_arq = f"digitalizado_{time.strftime('%H%M%S')}.pdf"
            path_final = os.path.join(USER_PICTURES, nome_arq)
            pdf.output(os.path.join(C_PATH, nome_arq))
            pdf.output(path_final)
            
            for p in self.paginas_scaneadas: 
                if os.path.exists(p): os.remove(p)
            
            self.after(0, lambda: self.sucesso_final(path_final))
        except Exception as e:
            self.after(0, lambda: [messagebox.showerror("Erro PDF", str(e)), self.reset_ui()])

    def sucesso_final(self, caminho):
        os.startfile(caminho)
        self.paginas_scaneadas = []
        self.atualizar_preview_completo()
        self.status.configure(text="PDF Criado com sucesso!")

    def reset_ui(self):
        self.btn_iniciar.configure(state="normal")
        self.status.configure(text="Pronto")

if __name__ == "__main__":
    app = ScannerApp()
    app.mainloop()