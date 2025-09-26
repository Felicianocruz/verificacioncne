import tkinter as tk
from tkinter import messagebox
from docx import Document
from docx2pdf import convert
import os
from datetime import datetime
import qrcode
import urllib.parse

# Ruta de la plantilla
PLANTILLA = "plantilla.docx"

def generar_credencial(nombre, dni, fecha):
    if not os.path.exists(PLANTILLA):
        messagebox.showerror("Error", f"No se encontró la plantilla: {PLANTILLA}")
        return None, None, None

    # Abrir plantilla
    doc = Document(PLANTILLA)

    # Reemplazar los marcadores en Word
    for p in doc.paragraphs:
        if "<<NOMBRE>>" in p.text:
            p.text = p.text.replace("<<NOMBRE>>", nombre)
        if "<<DNI>>" in p.text:
            p.text = p.text.replace("<<DNI>>", dni)
        if "<<FECHA>>" in p.text:
            p.text = p.text.replace("<<FECHA>>", fecha)

    # Evitar sobrescribir archivos existentes
    base_word = f"credencial_{dni}"
    word_file = base_word + ".docx"
    pdf_file = base_word + ".pdf"
    qr_file = base_word + "_QR.png"
    count = 1
    while os.path.exists(word_file) or os.path.exists(pdf_file) or os.path.exists(qr_file):
        word_file = f"{base_word}_{count}.docx"
        pdf_file = f"{base_word}_{count}.pdf"
        qr_file = f"{base_word}_{count}_QR.png"
        count += 1

    # Guardar Word
    doc.save(word_file)

    # Convertir a PDF
    convert(word_file, pdf_file)

    # Generar QR que lleve a búsqueda de Google
    query = urllib.parse.quote(nombre)  # Codificar nombre para URL
    url_busqueda = f"https://www.google.com/search?q={query}"
    qr_img = qrcode.make(url_busqueda)
    qr_img.save(qr_file)

    return word_file, pdf_file, qr_file

def generar_desde_gui():
    nombre = entry_nombre.get().strip()
    dni = entry_dni.get().strip()
    fecha = datetime.now().strftime("%d/%m/%Y")  # Fecha automática

    if not nombre or not dni:
        messagebox.showwarning("Campos vacíos", "Debes ingresar nombre y DNI.")
        return

    word_file, pdf_file, qr_file = generar_credencial(nombre, dni, fecha)
    if word_file and pdf_file and qr_file:
        messagebox.showinfo("Éxito", f"Credencial generada:\n{word_file}\n{pdf_file}\nQR: {qr_file}")

def limpiar_campos():
    entry_nombre.delete(0, tk.END)
    entry_dni.delete(0, tk.END)

# ----------------- INTERFAZ -----------------
ventana = tk.Tk()
ventana.title("Generador de Credencial CNE")
ventana.geometry("400x300")

tk.Label(ventana, text="Nombre completo:").pack(pady=5)
entry_nombre = tk.Entry(ventana, width=40)
entry_nombre.pack()

tk.Label(ventana, text="DNI:").pack(pady=5)
entry_dni = tk.Entry(ventana, width=40)
entry_dni.pack()

tk.Label(ventana, text=f"Fecha: {datetime.now().strftime('%d/%m/%Y')}").pack(pady=5)

tk.Button(ventana, text="Generar Credencial", command=generar_desde_gui, bg="#4CAF50", fg="white").pack(pady=10)
tk.Button(ventana, text="Limpiar Campos", command=limpiar_campos, bg="#f44336", fg="white").pack(pady=5)

ventana.mainloop()
