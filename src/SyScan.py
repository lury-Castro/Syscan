import io
import os
import time
import uuid
import tkinter as tk
from tkinter import messagebox, ttk

import win32com.client
from fpdf import FPDF
from PIL import Image

BASE_DIR = "scanners_SyScan"


# ================= WIA =================

def listar_drivers_wia():
    wia = win32com.client.Dispatch("WIA.DeviceManager")
    return [dev.Properties("Name").Value for dev in wia.DeviceInfos]


def conectar_driver(nome_driver):
    wia = win32com.client.Dispatch("WIA.DeviceManager")
    for dev in wia.DeviceInfos:
        if dev.Properties("Name").Value == nome_driver:
            return dev.Connect()
    raise Exception("Driver não encontrado!")


def criar_pasta_do_scanner(nome_driver):
    safe = "".join(c for c in nome_driver if c.isalnum() or c in "_ -")
    pasta = os.path.join(BASE_DIR, safe)
    pasta_temp = os.path.join(pasta, "temp")

    os.makedirs(pasta_temp, exist_ok=True)

    return pasta, pasta_temp



def scan_flatbed_to_file(nome_driver, pasta_temp, indice):
    device = conectar_driver(nome_driver)
    item = device.Items[0]

    try:
        item.Properties("Horizontal Resolution").Value = 300
        item.Properties("Vertical Resolution").Value = 300
    except:
        pass

    WIA_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
    image = item.Transfer(WIA_BMP)

    caminho = os.path.join(pasta_temp, f"pagina_{indice:03d}.png")
    img = Image.open(io.BytesIO(image.FileData.BinaryData))
    img.save(caminho, "PNG")

    return caminho


# ================= APP =================

class ScannerApp:
    def __init__(self, root):
        self.root = root
        root.title("Digitalizador - Mesa de Vidro")
        root.geometry("460x280")

        ttk.Label(root, text="Selecione o scanner:").pack(pady=5)

        self.combo = ttk.Combobox(root, state="readonly", width=50)
        self.combo.pack(pady=5)

        ttk.Button(root, text="Buscar Scanners", command=self.carregar).pack(pady=5)
        ttk.Button(root, text="Iniciar Digitalização", command=self.iniciar).pack(pady=15)

        self.status = ttk.Label(root, text="")
        self.status.pack(pady=5)

        self.nome_driver = None
        self.pasta = None
        self.contador = 1
        self.pasta_temp = None



    def carregar(self):
        drivers = listar_drivers_wia()
        if not drivers:
            messagebox.showerror("Erro", "Nenhum scanner SCANNER encontrado!")
            return
        self.combo["values"] = drivers
        self.combo.current(0)


    def iniciar(self):
        self.nome_driver = self.combo.get()
        if not self.nome_driver:
            messagebox.showwarning("Atenção", "Selecione um scanner!")
            return

        self.pasta,self.pasta_temp = criar_pasta_do_scanner(self.nome_driver)
        self.contador = 1
        self.status.config(text="Iniciando digitalização...")

        self.digitalizar_loop()


    def digitalizar_loop(self):
        self.status.config(text=f"Página {self.contador}")

        if not messagebox.askyesno(
            "Digitalizar",
            "Coloque a página no vidro e clique em SIM para digitalizar\n(Clique NÃO para cancelar)"
        ):
            self.finalizar_pdf()
            return

        try:
            scan_flatbed_to_file(
                self.nome_driver,
                self.pasta_temp,
                self.contador
            )
            messagebox.showinfo("OK", f"Página {self.contador} digitalizada!")
            self.contador += 1

        except Exception as e:
            messagebox.showerror("Erro", str(e))
            return

        if messagebox.askyesno("Continuar", "Deseja digitalizar outra página?"):
            self.digitalizar_loop()
        else:
            self.finalizar_pdf()


    def finalizar_pdf(self):
        imagens = sorted(
            [
                os.path.join(self.pasta_temp, f)
                for f in os.listdir(self.pasta_temp)
                if f.lower().endswith(".png")
            ]
        )

        if not imagens:
            messagebox.showinfo("Fim", "Nenhuma página digitalizada.")
            return

        output = os.path.join(
            self.pasta,
            f"scan_{time.strftime('%Y%m%d_%H%M%S')}.pdf"
        )

        pdf = FPDF(unit="mm", format="A4")

        for img_path in imagens:
            img = Image.open(img_path).convert("RGB")

            # Centralizar mantendo proporção
            w_img, h_img = img.size
            page_w, page_h = 210, 297

            ratio = min(page_w / w_img, page_h / h_img)
            new_w = w_img * ratio
            new_h = h_img * ratio

            x = (page_w - new_w) / 2
            y = (page_h - new_h) / 2

            temp = f"temp_{uuid.uuid4().hex}.jpg"
            img.save(temp, "JPEG", quality=70, optimize=True)

            pdf.add_page()
            pdf.image(temp, x=x, y=y, w=new_w, h=new_h)

            os.remove(temp)

        pdf.output(output)

        for img_path in imagens:
            try:
                os.remove(img_path)
            except Exception as e:
                print(f"Erro ao apagar {img_path}: {e}")

        messagebox.showinfo(
            "PDF Gerado",
            f"{len(imagens)} páginas geradas com sucesso!\n\n{output}"
        )

        try:
            os.startfile(output)
        except:
            pass
                # Apagar a pasta temp se só tem imagens
        try:
            arquivos = os.listdir(self.pasta_temp)

            apenas_imagens = all(
                f.lower().endswith((".png", ".jpg", ".jpeg"))
                for f in arquivos
            )

            if apenas_imagens:
                for f in arquivos:
                    os.remove(os.path.join(self.pasta_temp, f))

                os.rmdir(self.pasta_temp)
        except Exception as e:
            print(f"Erro ao limpar pasta temp: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScannerApp(root)
    root.mainloop()
