import io, os, time, uuid, tkinter as tk
from tkinter import messagebox, ttk
import win32com.client
from fpdf import FPDF
from PIL import Image

BASE_DIR = "scanners_SyScan"

def get_wia_devices():
    wia = win32com.client.Dispatch("WIA.DeviceManager")
    return {dev.Properties("Name").Value: dev for dev in wia.DeviceInfos}

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

class ScannerApp:
    def __init__(self, root):
        self.root = root
        root.title("ADAR - SyScaner")
        root.geometry("460x280")
        
        ttk.Label(root, text="Selecione o scanner:").pack(pady=5)
        self.combo = ttk.Combobox(root, state="readonly", width=50)
        self.combo.pack(pady=5)
        
        ttk.Button(root, text="Buscar Scanners", command=self.carregar).pack(pady=5)
        ttk.Button(root, text="Iniciar Digitalização", command=self.iniciar).pack(pady=15)
        self.status = ttk.Label(root, text="")
        self.status.pack()
        self.devices = {}

    def carregar(self):
        self.devices = get_wia_devices()
        self.combo["values"] = list(self.devices.keys())
        if self.devices: self.combo.current(0)
        else: messagebox.showerror("Erro", "Nenhum scanner encontrado!")

    def iniciar(self):
        nome = self.combo.get()
        if not nome: return messagebox.showwarning("Atenção", "Selecione um scanner!")
        
        safe_nome = "".join(c for c in nome if c.isalnum() or c in "_ -")
        self.pasta = os.path.join(BASE_DIR, safe_nome)
        self.pasta_temp = os.path.join(self.pasta, "temp")
        os.makedirs(self.pasta_temp, exist_ok=True)
        
        self.contador = 1
        self.digitalizar_loop()

    def digitalizar_loop(self):
        if not messagebox.askyesno("ADAR - SyScaner", f"Página {self.contador}: Coloque no vidro e clique SIM."):
            return self.finalizar_pdf()

        try:
            scan_to_file(self.devices[self.combo.get()], self.pasta_temp, self.contador)
            self.contador += 1
            if messagebox.askyesno("ADAR - SyScaner", "Deseja digitalizar outra página?"):
                self.digitalizar_loop()
            else: self.finalizar_pdf()
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def finalizar_pdf(self):
        imgs = [os.path.join(self.pasta_temp, f) for f in sorted(os.listdir(self.pasta_temp)) if f.endswith(".png")]
        if not imgs: return messagebox.showinfo("Aviso", "Nada digitalizado.")

        pdf = FPDF()
        for p in imgs:
            with Image.open(p).convert("RGB") as img:
                w, h = img.size
                ratio = min(210/w, 297/h)
                nw, nh = w*ratio, h*ratio
                temp_jpg = f"temp_{uuid.uuid4().hex}.jpg"
                img.save(temp_jpg, "JPEG", quality=70)
                pdf.add_page()
                pdf.image(temp_jpg, (210-nw)/2, (297-nh)/2, nw, nh)
                os.remove(temp_jpg)

        out = os.path.join(self.pasta, f"scan_{time.strftime('%Y%m%d_%H%M%S')}.pdf")
        pdf.output(out)
        
        for p in imgs: os.remove(p)
        try: os.rmdir(self.pasta_temp)
        except: pass
        
        messagebox.showinfo("ADAR - SyScaner", f"PDF gerado: {len(imgs)} páginas")
        os.startfile(out)

if __name__ == "__main__":
    root = tk.Tk()
    ScannerApp(root)
    root.mainloop()