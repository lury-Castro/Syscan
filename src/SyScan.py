import io, os, time, uuid, tkinter as tk
from tkinter import messagebox
import customtkinter as ctk 
import win32com.client
from fpdf import FPDF
from PIL import Image

# --- Configurações Visuais ---
ctk.set_appearance_mode("Light")  # Adaptativo (Claro/Escuro)
ctk.set_default_color_theme("blue") 

BASE_DIR = "scanners_SyScan"

def get_wia_devices():
    try:
        wia = win32com.client.Dispatch("WIA.DeviceManager")
        return {dev.Properties("Name").Value: dev for dev in wia.DeviceInfos}
    except:
        return {}

def scan_to_file(device_info, pasta_temp, indice):
    dev = device_info.Connect()
    item = dev.Items[0]
    for prop in ["Horizontal Resolution", "Vertical Resolution"]:
        try: item.Properties(prop).Value = 300
        except: pass
    
    image = item.Transfer("{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")
    caminho = os.path.join(pasta_temp, f"pg_{indice:03d}.png")
    Image.open(io.BytesIO(image.FileData.BinaryData)).save(caminho, "PNG")
    return caminho

class ScannerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações da Janela principal
        self.title("ADAR - SyScaner")
        self.geometry("650x500")
        
        # Caminho absoluto para evitar erros com imagens
        self.diretorio_script = os.path.dirname(os.path.abspath(__file__))
        caminho_nome = os.path.join(self.diretorio_script, "adar.png")
        caminho_logo = os.path.join(self.diretorio_script, "logo.ico")

        # Define o ícone da barra de tarefas (canto superior esquerdo)
        if os.path.exists(caminho_logo):
            try:
                self.after(200, lambda: self.iconbitmap(caminho_logo))
            except:
                pass

        # --- FRAME DO LOGOTIPO (TOPO) ---
        self.logo_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.logo_frame.pack(pady=25)
        
        if os.path.exists(caminho_nome):
            try:
                img_puro = Image.open(caminho_nome)
                # Criando imagem compatível com CustomTkinter
                self.logo_img = ctk.CTkImage(light_image=img_puro, 
                                            dark_image=img_puro, 
                                            size=(200,50)) # Ajuste aqui o tamanho do logo
                self.logo_label = ctk.CTkLabel(self.logo_frame, image=self.logo_img, text="")
                self.logo_label.pack()
            except Exception as e:
                self.logo_label = ctk.CTkLabel(self.logo_frame, text="ADAR SYSCANER", font=("Roboto", 24, "bold"))
                self.logo_label.pack()
        else:
            self.logo_label = ctk.CTkLabel(self.logo_frame, text="ADAR SYSCANER", font=("Roboto", 24, "bold"))
            self.logo_label.pack()

        # --- CONTEÚDO CENTRAL ---
        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.pack(pady=10, padx=30, fill="both", expand=True)

        self.label_scanner = ctk.CTkLabel(self.main_frame, text="Selecione o Scanner:", font=("Roboto", 14, "bold"))
        self.label_scanner.pack(pady=(20, 5))

        self.combo = ctk.CTkComboBox(self.main_frame, width=400, state="readonly")
        self.combo.pack(pady=5)
        self.combo.set("Clique em Buscar Scanner")

        self.btn_buscar = ctk.CTkButton(self.main_frame, text="Buscar Scanner", 
                                        command=self.carregar,
                                        fg_color="#161718",
                                        hover_color="#555658",
                                        border_width=1)
        self.btn_buscar.pack(pady=10)

        self.btn_iniciar = ctk.CTkButton(self.main_frame, text="INICIAR DIGITALIZAÇÃO", 
                                         command=self.iniciar, 
                                         height=50, width=300,
                                         font=("Roboto", 16, "bold"),
                                         fg_color="#161718", hover_color="#555658")
        self.btn_iniciar.pack(pady=30)

        # --- BARRA DE STATUS (RODAPÉ) ---
        self.status = ctk.CTkLabel(self, text="Pronto para iniciar", font=("Roboto", 11), text_color="#555658")
        self.status.pack(side="bottom", pady=5)
        
        self.devices = {}

    def carregar(self):
        self.devices = get_wia_devices()
        nomes = list(self.devices.keys())
        if nomes:
            self.combo.configure(values=nomes)
            self.combo.set(nomes[0])
            self.status.configure(text=f"Sucesso: {len(nomes)} scanner(s) encontrado(s)",
                                  text_color="#2e4bcc")
        else:
            self.status.configure(text="Nenhum scanner detectado", text_color="#e74c3c")
            messagebox.showerror("Erro", "Nenhum scanner encontrado! Verifique a conexão USB.")

    def iniciar(self):
        nome = self.combo.get()
        if nome in ["Clique em Buscar Scanner", ""]: 
            return messagebox.showwarning("Atenção", "Por favor, busque e selecione um scanner primeiro!")
        
        safe_nome = "".join(c for c in nome if c.isalnum() or c in "_ -")
        self.pasta = os.path.join(BASE_DIR, safe_nome)
        self.pasta_temp = os.path.join(self.pasta, "temp")
        os.makedirs(self.pasta_temp, exist_ok=True)
        
        self.contador = 1
        self.digitalizar_loop()

    def digitalizar_loop(self):
        if not messagebox.askyesno("ADAR-SyScaner", f"Página {self.contador}:\nPosicione o documento e clique em SIM para escanear."):
            return self.finalizar_pdf()

        try:
            self.status.configure(text=f"Processando página {self.contador}...", text_color="#3498db")
            self.update() 
            
            scan_to_file(self.devices[self.combo.get()], self.pasta_temp, self.contador)
            self.contador += 1
            
            if messagebox.askyesno("ADAR-SyScaner", "Deseja digitalizar a PRÓXIMA página?"):
                self.digitalizar_loop()
            else:
                self.finalizar_pdf()
        except Exception as e:
            self.status.configure(text="Erro durante a digitalização", text_color="#e74c3c")
            messagebox.showerror("Erro de Hardware", f"Falha ao comunicar com o scanner:\n{str(e)}")

    def finalizar_pdf(self):
        self.status.configure(text="Gerando arquivo PDF...", text_color="#f39c12")
        self.update()
        
        imgs = [os.path.join(self.pasta_temp, f) for f in sorted(os.listdir(self.pasta_temp)) if f.endswith(".png")]
        if not imgs: 
            self.status.configure(text="Operação cancelada", text_color="gray")
            return

        pdf = FPDF()
        for p in imgs:
            with Image.open(p).convert("RGB") as img:
                w, h = img.size
                ratio = min(210/w, 297/h)
                nw, nh = w*ratio, h*ratio
                temp_jpg = f"temp_{uuid.uuid4().hex}.jpg"
                img.save(temp_jpg, "JPEG", quality=75)
                pdf.add_page()
                pdf.image(temp_jpg, (210-nw)/2, (297-nh)/2, nw, nh)
                os.remove(temp_jpg)

        out = os.path.join(self.pasta, f"scan_{time.strftime('%Y%m%d_%H%M%S')}.pdf")
        pdf.output(out)
        
        # Limpeza
        for p in imgs: os.remove(p)
        try: os.rmdir(self.pasta_temp)
        except: pass
        
        self.status.configure(text="PDF gerado com sucesso!", text_color="#2ecc71")
        messagebox.showinfo("ADAR-SyScaner", f"Arquivo PDF criado com {len(imgs)} páginas.")
        os.startfile(out)

if __name__ == "__main__":
    app = ScannerApp()
    app.mainloop()