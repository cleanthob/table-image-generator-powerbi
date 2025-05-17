import os
import base64
import pandas as pd
import ttkbootstrap as ttk
from tkinter import messagebox, filedialog


def convert_img_to_b64(img_path):
    ext = os.path.splitext(img_path)[-1].lower()
    if ext in [".jpg", ".jpeg"]:
        mime_type = "image/jpeg"
    elif ext == ".png":
        mime_type = "image/png"
    else:
        mime_type = "image/unknown"
    with open(img_path, "rb") as image_file:
        base64_string = base64.b64encode(image_file.read()).decode("utf-8")
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
            img_b64 = convert_img_to_b64(path_img)
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


app = ttk.Window(
    title="Conversor de imagens para Base64", themename="cyborg", size=(800, 280)
)
app.resizable(False, False)

pasta_var = ttk.StringVar()

ttk.Label(
    app,
    text="Conversor de Imagens para Base64 (Power BI)",
    font=("Segoe UI", 18, "bold"),
).pack(pady=15)

frame = ttk.Frame(app)
frame.pack(padx=20, fill="x")

entry_pasta = ttk.Entry(frame, textvariable=pasta_var)
entry_pasta.pack(side="left", fill="x", expand=True, pady=10)

btn_selecionar = ttk.Button(
    frame, text="Selecionar Pasta", command=selecionar_pasta, bootstyle="primary"
)
btn_selecionar.pack(side="left", padx=10, pady=10)

btn_exportar = ttk.Button(
    app, text="Exportar para Excel", command=exportar_excel, bootstyle="success"
)
btn_exportar.pack(pady=20)

label_status = ttk.Label(
    app, text="Nenhuma pasta selecionada", font=("Segoe UI", 10), foreground="red"
)
label_status.pack()

app.mainloop()
