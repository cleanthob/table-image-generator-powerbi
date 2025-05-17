import os
import base64
import pandas as pd
import ttkbootstrap as ttk
from tkinter import messagebox, filedialog
from PIL import Image
import io


def convert_img_to_b64_redimensionada(img_path, max_size=(200, 200)):
    ext = os.path.splitext(img_path)[-1].lower()
    if ext in [".jpg", ".jpeg"]:
        mime_type = "image/jpeg"
    elif ext == ".png":
        mime_type = "image/png"
    else:
        mime_type = "image/unknown"
    img = Image.open(img_path)
    img.thumbnail(max_size)
    buffered = io.BytesIO()
    format_save = "JPEG" if mime_type == "image/jpeg" else "PNG"
    img.save(buffered, format=format_save)
    base64_string = base64.b64encode(buffered.getvalue()).decode("utf-8")
    return f"data:{mime_type};base64,{base64_string}"


def selecionar_pasta():
    pasta = filedialog.askdirectory(title="Selecione a pasta com imagens")
    if pasta:
        pasta_var.set(pasta)
        label_status.config(text=f"Pasta selecionada: {pasta}", foreground="green")


def exportar_excel():
    pasta = pasta_var.get()
    if not pasta:
        messagebox.showerror("Erro", "Nenhuma pasta foi selecionada.")
        return
    data = []
    for img in os.listdir(pasta):
        path_img = os.path.join(pasta, img)
        if os.path.isfile(path_img) and img.lower().endswith((".png", ".jpg", ".jpeg")):
            img_name = os.path.splitext(img)[0]
            img_b64 = convert_img_to_b64_redimensionada(path_img)
            data.append({"name": img_name, "b64": img_b64})
    if not data:
        messagebox.showerror(
            "Erro", "Nenhuma imagem válida (.png, .jpg, .jpeg) encontrada na pasta."
        )
        return
    df = pd.DataFrame(data)
    caminho_salvar = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")],
        title="Salvar como",
    )
    if caminho_salvar:
        df.to_excel(caminho_salvar, index=False)
        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso:\n{caminho_salvar}")
    else:
        messagebox.showwarning("Cancelado", "Operação de salvamento cancelada.")


width = 600
height = 280

app = ttk.Window(
    title="Conversor de imagens para Base64", themename="morph", size=(width, height)
)
app.resizable(True, True)
app.update_idletasks()

x = (app.winfo_screenwidth() // 2) - (width // 2)
y = (app.winfo_screenheight() // 2) - (height // 2)
app.geometry(f"{width}x{height}+{x}+{y}")

pasta_var = ttk.StringVar()

frame = ttk.Frame(app)
frame.pack(padx=20, pady=10, fill="x")

entry_pasta = ttk.Entry(frame, textvariable=pasta_var)
entry_pasta.pack(side="left", fill="x", expand=True)

btn_selecionar = ttk.Button(
    frame, text="Selecionar Pasta", command=selecionar_pasta, bootstyle="primary"
)
btn_selecionar.pack(side="left", padx=10)

btn_exportar = ttk.Button(
    app, text="Exportar para Excel", command=exportar_excel, bootstyle="success"
)
btn_exportar.pack(pady=10, padx=20, fill="x")

label_status = ttk.Label(
    app, text="Nenhuma pasta selecionada", font=("Segoe UI", 10), foreground="red"
)
label_status.pack(padx=20, pady=10, fill="x")

app.mainloop()
