import pymupdf
import os
import fitz
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from datetime import datetime
import subprocess
import sys
import json
import shutil

def guardar_ultima_carpeta(carpeta):
    """Guarda la última carpeta seleccionada en un archivo JSON."""
    with open("ultima_carpeta.json", "w") as f:
        json.dump({"ultima_carpeta": carpeta}, f)

def cargar_ultima_carpeta():
    """Carga la última carpeta seleccionada desde el archivo JSON."""
    if os.path.exists("ultima_carpeta.json"):
        with open("ultima_carpeta.json", "r") as f:
            datos = json.load(f)
            return datos.get("ultima_carpeta", os.path.dirname(os.path.abspath(__file__)))
    return os.path.dirname(os.path.abspath(__file__))

def guardar_cache(cache):
    """Guarda el cache de firmas en un archivo JSON."""
    with open("cache_firmas.json", "w") as f:
        json.dump(cache, f)

def cargar_cache():
    """Carga el cache de firmas desde un archivo JSON."""
    if os.path.exists("cache_firmas.json"):
        with open("cache_firmas.json", "r") as f:
            return json.load(f)
    return {}

global selected_folder
selected_folder = cargar_ultima_carpeta()
cache_firmas = cargar_cache()

def verificar_firma(pdf_path):
    """Verifica si un archivo PDF está firmado digitalmente y obtiene el firmante."""
    if pdf_path in cache_firmas:
        return cache_firmas[pdf_path]
    
    try:
        doc = fitz.open(pdf_path)
        for xref in range(1, doc.xref_length()):
            obj = doc.xref_object(xref)
            if "/Type /Sig" in obj:
                firmante = "Firmante desconocido"
                if " /Name" in obj:
                    firmante = obj.split("/Name (")[1].split(")")[0].strip()
                cache_firmas[pdf_path] = (True, firmante)
                guardar_cache(cache_firmas)
                return True, firmante
    except Exception as e:
        print(f"Error al verificar {pdf_path}: {e}")
    
    cache_firmas[pdf_path] = (False, "N/A")
    guardar_cache(cache_firmas)
    return False, "N/A"

def listar_pdfs(filtro="todos"):
    """Actualiza la lista de PDFs en la interfaz."""
    for row in tree.get_children():
        tree.delete(row)
    
    if not selected_folder:
        return
    
    global pdf_data
    pdf_data = []
    
    try:
        for archivo in os.listdir(selected_folder):
            if archivo.endswith(".pdf"):
                ruta = os.path.join(selected_folder, archivo).replace("/", "\\")
                firmado, firmante = verificar_firma(ruta)
                fecha_mod = datetime.fromtimestamp(os.path.getmtime(ruta)).strftime('%Y-%m-%d %H:%M:%S')
                if filtro == "firmados" and not firmado:
                    continue
                if filtro == "no firmados" and firmado:
                    continue
                pdf_data.append((archivo, fecha_mod, "Sí" if firmado else "No", firmante, ruta))
    except Exception as e:
        print(f"Error al listar archivos en la carpeta: {e}")
    
    actualizar_treeview()

def actualizar_treeview():
    """Llena el TreeView con los datos ordenados."""
    for row in tree.get_children():
        tree.delete(row)
    
    for archivo, fecha_mod, firmado, firmante, ruta in pdf_data:
        color = "success" if firmado == "Sí" else "danger"
        tree.insert("", "end", values=(archivo, fecha_mod, firmado, firmante, ruta), tags=(color,))
        tree.tag_configure("success", background="#d4edda", foreground="#000000")
        tree.tag_configure("danger", background="#f8d7da", foreground="#000000")

def abrir_pdf(ruta):
    """Abre el archivo PDF con el programa predeterminado del sistema."""
    try:
        ruta = ruta.replace("/", "\\")  # Reemplaza las barras diagonales con barras invertidas
        if os.path.exists(ruta):
            if sys.platform.startswith("win"):
                os.startfile(ruta)
            elif sys.platform.startswith("darwin"):
                subprocess.run(["open", ruta])
            else:
                subprocess.run(["xdg-open", ruta])
    except Exception as e:
        print(f"No se pudo abrir el archivo: {e}")

def ordenar_por(columna):
    """Ordena la lista de PDFs en orden ascendente o descendente al alternar clics."""
    global orden_actual
    orden_actual[columna] = not orden_actual.get(columna, False)
    reverse_order = orden_actual[columna]
    
    if columna == "Fecha":
        pdf_data.sort(key=lambda x: datetime.strptime(x[1], '%Y-%m-%d %H:%M:%S'), reverse=reverse_order)
    elif columna == "Nombre":
        pdf_data.sort(key=lambda x: x[0], reverse=reverse_order)
    elif columna == "Firmado":
        pdf_data.sort(key=lambda x: x[2], reverse=reverse_order)
    
    actualizar_treeview()

def seleccionar_carpeta():
    """Permite seleccionar una carpeta y guarda la última seleccionada."""
    global selected_folder
    folder = filedialog.askdirectory()
    if folder:
        selected_folder = folder.replace("/", "\\")
        guardar_ultima_carpeta(folder)
        listar_pdfs()

def abrir_archivo():
    """Abre el archivo seleccionado en la tabla."""
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item[0], 'values')
        abrir_pdf(item[4])

def abrir_archivo_doble_click(event):
    """Abre el archivo PDF seleccionado en la tabla cuando se hace doble clic."""
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item[0], 'values')
        abrir_pdf(item[4])

def ver_todos():
    """Muestra todos los archivos en la lista."""
    listar_pdfs(filtro="todos")

def ver_firmados():
    """Muestra solo los archivos firmados en la lista."""
    listar_pdfs(filtro="firmados")

def ver_no_firmados():
    """Muestra solo los archivos no firmados en la lista."""
    listar_pdfs(filtro="no firmados")

def cargar_nuevos_archivos():
    """Permite cargar nuevos archivos PDF en la carpeta seleccionada."""
    global selected_folder
    if selected_folder:
        archivos = filedialog.askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
        if archivos:
            for archivo in archivos:
                nombre_archivo = os.path.basename(archivo)
                nueva_ruta = os.path.join(selected_folder, nombre_archivo).replace("/", "\\")
                shutil.copy2(archivo, nueva_ruta)
            listar_pdfs()

# Configuración de la interfaz gráfica con ttkbootstrap
root = ttkb.Window(themename="flatly")
style = ttkb.Style()

root.title("Ordenes de pagos - PDF")
root.geometry("800x600")

# Crear la barra de menú tipo hamburguesa
menubar = ttkb.Menu(root)
opciones_menu = ttkb.Menu(menubar, tearoff=0)
opciones_menu.add_command(label="Ver todos", command=ver_todos)
opciones_menu.add_command(label="Ver firmados", command=ver_firmados)
opciones_menu.add_command(label="Ver no firmados", command=ver_no_firmados)
menubar.add_cascade(label="Opciones", menu=opciones_menu)
root.config(menu=menubar)

global orden_actual
orden_actual = {}

frame = ttkb.Frame(root)
frame.pack(pady=10, padx=10)  # Agregar padding interno al frame

btn_seleccionar = ttkb.Button(frame, text="Seleccionar Carpeta", command=seleccionar_carpeta, bootstyle="success-outline")
btn_seleccionar.pack(side="left", padx=5)

btn_actualizar = ttkb.Button(frame, text="Actualizar", command=ver_todos, bootstyle="success-outline")
btn_actualizar.pack(side="left", padx=5)

btn_abrir = ttkb.Button(frame, text="Abrir Archivo", command=abrir_archivo, bootstyle="success-outline")
btn_abrir.pack(side="left", padx=5)

btn_cargar = ttkb.Button(frame, text="+", command=cargar_nuevos_archivos, bootstyle="success-outline")
btn_cargar.pack(side="left", padx=5)

columns = ("Nombre", "Fecha", "Firmado", "Firmante", "Ruta")
tree = ttkb.Treeview(root, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col, command=lambda c=col: ordenar_por(c))
    tree.column(col, width=150)
tree.pack(expand=True, fill="both", padx=10, pady=10)  # Agregar padding interno al treeview

# Vincular el evento de doble clic al TreeView
tree.bind("<Double-1>", abrir_archivo_doble_click)

# Cargar automáticamente la última carpeta seleccionada
listar_pdfs()

root.mainloop()
