import os
import base64
import pandas as pd
import ttkbootstrap as ttk
from tkinter import messagebox, filedialog, Text, Scrollbar, HORIZONTAL, BOTTOM, X
from PIL import Image
import io


def convert_img_to_b64_resized(img_path, max_size=(200, 200)):
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


def select_folder():
    folder = filedialog.askdirectory(title="Selecione a pasta com imagens")
    if folder:
        folder_var.set(folder)
        update_status(f"Pasta selecionada: {folder}", "green")


def export_to_excel():
    folder = folder_var.get()
    if not folder:
        messagebox.showerror("Erro", "Nenhuma pasta foi selecionada.")
        return
    data = []
    for img in os.listdir(folder):
        path_img = os.path.join(folder, img)
        if os.path.isfile(path_img) and img.lower().endswith((".png", ".jpg", ".jpeg")):
            img_name = os.path.splitext(img)[0]
            img_b64 = convert_img_to_b64_resized(path_img)
            data.append({"name": img_name, "b64": img_b64})
    if not data:
        messagebox.showerror(
            "Erro", "Nenhuma imagem válida (.png, .jpg, .jpeg) encontrada na pasta."
        )
        return
    df = pd.DataFrame(data)
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")],
        title="Salvar como",
    )
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso:\n{save_path}")
    else:
        messagebox.showwarning("Cancelado", "Operação de salvamento cancelada.")


def update_status(text, color):
    status_text.config(state="normal")
    status_text.delete("1.0", "end")
    status_text.insert("end", text)
    status_text.config(state="disabled", foreground=color)


width = 800
height = 280

app = ttk.Window(
    title="Conversor de imagens para Base64", themename="lumen", size=(width, height)
)
app.resizable(True, True)
app.update_idletasks()

x = (app.winfo_screenwidth() // 2) - (width // 2)
y = (app.winfo_screenheight() // 2) - (height // 2)
app.geometry(f"{width}x{height}+{x}+{y}")

folder_var = ttk.StringVar()

frame = ttk.Frame(app)
frame.pack(padx=20, pady=20, fill="x", anchor="center")

entry_folder = ttk.Entry(frame, textvariable=folder_var)
entry_folder.pack(side="left", fill="x", expand=True)

btn_select = ttk.Button(
    frame,
    text="Selecionar Pasta",
    command=select_folder,
    bootstyle="primary",
    width=15,
)
btn_select.pack(side="left", padx=(10, 0))

btn_export = ttk.Button(
    app,
    text="Exportar para Excel",
    command=export_to_excel,
    bootstyle="success",
    width=20,
)
btn_export.pack(pady=10)


status_frame = ttk.Frame(app)
status_frame.pack(padx=20, pady=10, fill="x")

status_text = Text(
    status_frame,
    height=1,
    wrap="none",
    font=("Segoe UI", 10),
    foreground="red",
    state="disabled",
)
status_text.pack(side="top", fill="x", expand=True)

scroll_x = Scrollbar(status_frame, orient=HORIZONTAL, command=status_text.xview)
scroll_x.pack(side=BOTTOM, fill=X)
status_text.config(xscrollcommand=scroll_x.set)

update_status("Nenhuma pasta selecionada", "red")

app.mainloop()
