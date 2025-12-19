import io, os, time, uuid
import customtkinter as ctk
import win32com.client
from tkinter import messagebox
from fpdf import FPDF
from PIL import Image

# --- Configurações Globais ---
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")
BASE_DIR = "scanners_SyScan"

def get_wia_devices():
    """Retorna dicionário de scanners disponíveis."""
    try:
        wia = win32com.client.Dispatch("WIA.DeviceManager")
        return {dev.Properties("Name").Value: dev for dev in wia.DeviceInfos}
    except Exception: return {}

def scan_to_file(device_info, pasta_temp, indice):
    """Executa a digitalização física."""
    dev = device_info.Connect()
    item = dev.Items[0]
    for prop in ["Horizontal Resolution", "Vertical Resolution"]:
        try: item.Properties(prop).Value = 300
        except Exception: pass
    
    image = item.Transfer("{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")
    caminho = os.path.join(pasta_temp, f"pg_{indice:03d}.png")
    Image.open(io.BytesIO(image.FileData.BinaryData)).save(caminho, "PNG")
    return caminho

class ScannerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ADAR - SyScaner")
        self.geometry("600x350")
        self.devices = {}
        
        # Caminhos e Recursos
        self.path_root = os.path.dirname(os.path.abspath(__file__))
        self._setup_ui()
        self._load_icons()
        self.update_status("Buscando scanners conectados...", "#3498db")
        self.after(500, self.carregar)

    def _load_icons(self):
        """Carrega ícone da janela e logo com tratamento de erro simplificado."""
        icon_path = os.path.join(self.path_root, "logo.ico")
        img_path = os.path.join(self.path_root, "adar.png")
        
        if os.path.exists(icon_path):
            self.after(200, lambda: self.iconbitmap(icon_path))
            
        if os.path.exists(img_path):
            img = ctk.CTkImage(Image.open(img_path), size=(200, 50))
            self.logo_label.configure(image=img, text="")

    def _setup_ui(self):
        """Constrói a interface de forma scaneável."""
        # Logo Area
        self.logo_label = ctk.CTkLabel(self, text="ADAR SYSCANER", font=("Roboto", 24, "bold"))
        self.logo_label.pack(pady=20)

        # Main Container
        self.container = ctk.CTkFrame(self, corner_radius=15)
        self.container.pack(pady=10, padx=30, fill="both", expand=True)

        ctk.CTkLabel(self.container, text="Selecione o Scanner:", font=("Roboto", 14, "bold")).pack(pady=(20, 5))
        
        self.combo = ctk.CTkComboBox(self.container, width=350, state="readonly")
        self.combo.set("Clique em Buscar Scanner")
        self.combo.pack(pady=5)

        # Botões usando um padrão comum
        btn_args = {"master": self.container, "fg_color": "#161718", "hover_color": "#555658"}
        

        ctk.CTkButton(**btn_args, text="INICIAR DIGITALIZAÇÃO", height=50, width=280, 
                      font=("Roboto", 16, "bold"), command=self.iniciar).pack(pady=30)

        self.status = ctk.CTkLabel(self, text="Pronto para iniciar", font=("Roboto", 11), text_color="#555658")
        self.status.pack(side="bottom", pady=10)

    def update_status(self, text, color="#555658"):
        """Centraliza a atualização de status."""
        self.status.configure(text=text, text_color=color)
        self.update()

    def carregar(self):
        self.devices = get_wia_devices()
        if self.devices:
            nomes = list(self.devices.keys())
            self.combo.configure(values=nomes)
            self.combo.set(nomes[0])
            self.update_status(f"Sucesso: {len(nomes)} scanner(s) encontrado(s)", "#2e4bcc")
        else:
            messagebox.showerror("Erro", "Nenhum scanner encontrado!")

    def iniciar(self):
        nome = self.combo.get()
        if nome in ["Clique em Buscar Scanner", ""]:
            return messagebox.showwarning("Atenção", "Selecione um scanner!")
        
        # Preparação de pasta
        safe_nome = "".join(c for c in nome if c.isalnum() or c in "_ -")
        self.pasta = os.path.join(BASE_DIR, safe_nome)
        self.pasta_temp = os.path.join(self.pasta, "temp")
        os.makedirs(self.pasta_temp, exist_ok=True)
        
        self.contador = 1
        self.digitalizar_loop()

    def digitalizar_loop(self):
        msg = f"Página {self.contador}:\nPosicione o documento e clique em SIM."
        if not messagebox.askyesno("ADAR-SyScaner", msg):
            return self.finalizar_pdf()

        try:
            
            self.update_status(f"AGUARDE: Digitalizando página {self.contador}...", "#3498db")
            
            self.update_idletasks() 
            
            scan_to_file(self.devices[self.combo.get()], self.pasta_temp, self.contador)
            self.contador += 1
            
            if messagebox.askyesno("ADAR-SyScaner", "Deseja digitalizar a PRÓXIMA página?"):
                self.digitalizar_loop()
            else:
                self.finalizar_pdf()
        except Exception as e:
            self.update_status("Erro na digitalização", "#e74c3c")
            messagebox.showerror("Erro de Hardware", f"Falha: {e}")

    def finalizar_pdf(self):
        self.update_status("Gerando PDF...", "#f39c12")
        imgs = [os.path.join(self.pasta_temp, f) for f in sorted(os.listdir(self.pasta_temp)) if f.endswith(".png")]
        
        if not imgs:
            self.update_status("Operação cancelada", "gray")
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
        
        # Limpeza rápida
        for p in imgs: os.remove(p)
        if os.path.exists(self.pasta_temp): os.rmdir(self.pasta_temp)
        
        self.update_status("PDF gerado com sucesso!", "#2ecc71")
        messagebox.showinfo("ADAR - SyScan", f"PDF criado com {len(imgs)} páginas.")
        os.startfile(out)

if __name__ == "__main__":
    ScannerApp().mainloop()