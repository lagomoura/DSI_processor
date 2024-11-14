from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import io
import re
import os
import time
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import threading
import shutil
import sys
from pathlib import Path

if getattr(sys, 'frozen', False):
    # El archivo está empaquetado con PyInstaller
    base_directory = os.path.dirname(sys.executable)
else:
    # El archivo se está ejecutando directamente desde el script
    base_directory = os.path.dirname(os.path.abspath(__file__))

# Define el directorio a observar
WATCH_DIRECTORY = os.path.join(base_directory, "input_pdf")
SIGNATURE_IMAGE = os.path.join(base_directory, "signature", "firma.JPEG")

if not os.path.exists(WATCH_DIRECTORY):
    print(f"Error: El directorio {WATCH_DIRECTORY} no existe o no es accesible.")
# Finaliza el programa si la carpeta no existe

def extract_text_from_all_pages(pdf_file_path):
    reader = PdfReader(pdf_file_path)
    all_text = [page.extract_text() for page in reader.pages]
    return all_text

def extract_cuit_pdf(page):
    # Leer el PDF
    text = page.extract_text()
    match = re.search(r'CUIT\s*-\s*(\d+\s*\d*)', text)

    if match:
        cuit = match.group(1)
        cuit_clean = cuit.replace(" ", "")

        return f"{cuit_clean}"

    else:
        return None

def read_nro_guia(all_text, page_number):
    text = all_text[page_number].split("\n")
    if text:
        doc_transp = text[10].split()[0]
        return doc_transp
    else:
        print("Error leyendo doc_transpot", text)
        return None

def split_pdf_add_img(pdf_file_path, image_path, log_output):

    try:
        time.sleep(1)
        reader = PdfReader(pdf_file_path)

        # Leer la firma
        img = Image.open(image_path)

        # Variables para el proceso de agrupacion de paginas por CUIT
        current_cuit_code = None
        current_writer = None
        current_output_folder = None
        pages_buffer = []
        start_page_num = 0
        
        all_text = extract_text_from_all_pages(pdf_file_path)
        
        # Ignorando la ultima hoja (resumen de DSI)
        total_pages = len(reader.pages) - 1
        
        for page_num in range(total_pages):

            page = reader.pages[page_num]
            cuit_code = extract_cuit_pdf(page)
            
            if cuit_code is None:
                log_output.insert(tk.END, f"Error: No se pudo leer el CUIT de la página {page_num + 1}\n")
                log_output.see(tk.END)
                continue

        # Si encontramos un nuevo CUIT, guardamos el bloque actual (si existe)
            if current_cuit_code is not None and cuit_code != current_cuit_code:
                # Agregar la firma a la última página del bloque antes de guardar
                pages_buffer[-1] = add_signature_to_page(pages_buffer[-1], image_path, img)
                save_pdf_block(current_writer, pages_buffer, current_output_folder, log_output, all_text, start_page_num)
                pages_buffer = []
                start_page_num = page_num

                
            # Carpeta de salida para este CUIT
            p_directory = "\\\\10.55.55.9\\particulares"
            cuit_folder = os.path.join(p_directory, cuit_code)
            
            if not os.path.exists(cuit_folder):
                os.makedirs(cuit_folder, exist_ok=True)
                log_output.insert(tk.END, f"Creando carpeta {cuit_folder}\n")
            else:
                log_output.insert(tk.END, f"La carpeta {cuit_folder} ya existe\n")

            current_cuit_code = cuit_code
            current_output_folder = os.path.join(p_directory, current_cuit_code)
            os.makedirs(current_output_folder, exist_ok=True)
            
            pages_buffer.append(page)

        # Guardar el último bloque al finalizar el bucle
        if pages_buffer:
            save_pdf_block(current_writer, pages_buffer, current_output_folder, log_output, all_text, total_pages - len(pages_buffer))

        log_output.insert(tk.END, f"Procesamiento finalizado\n")
        log_output.see(tk.END)
        mover_pdf_a_procesados(pdf_file_path, log_output)

    except PermissionError as e:
        log_output.insert(tk.END, f"Error: No se pudo acceder al archivo {pdf_file_path} debido a permisos. {e}\n")
        log_output.see(tk.END)
        time.sleep(2)  # Espera antes de intentar de nuevo
        split_pdf_add_img(pdf_file_path, image_path, log_output)  # Intentar nuevamente


def add_signature_to_page(page, image_path, img):
    # Crear un archivo en blanco con las mismas configuraciones de la página original
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # Obteniendo tamano de la pagina para centrar la firma
    page_width, page_height = letter
    img_width, img_height = img.size

    # Calculando x e y para centrar la firma
    x = (page_width - img_width) / 13
    y = (page_height - img_height) / 13

    # Dibujando la firma en el centro de la pagina
    can.drawImage(image_path, x, y, width=img_width, height=img_height)
    can.save()

    # Moviendo el archivo en blanco a la pagina extraida
    packet.seek(0)
    new_pdf = PdfReader(packet)
    page.merge_page(new_pdf.pages[0])

    return page


def save_pdf_block(writer, pages_buffer, output_folder, log_output, all_text, start_page_num):    

    for page_num, page in enumerate(pages_buffer):
        writer = PdfWriter()
        page_index = start_page_num + page_num
        
        # Guardar el bloque de páginas con el CUIT actual
        nombre_archivo = read_nro_guia(all_text, page_index)
        if not nombre_archivo:
            nombre_archivo = f"Pagina_{page_index + 1}"
        output_pdf_path = os.path.join(output_folder, f"{nombre_archivo}.pdf")
        
        # Verificar que la ruta del archivo sea válida
        if not os.path.isdir(output_folder):
            log_output.insert(tk.END, f"Error: El directorio de salida no existe: {output_folder}\n")
            log_output.see(tk.END)
            return
        
        try:
            writer.add_page(page)

            with open(output_pdf_path, "wb") as output_pdf:
                writer.write(output_pdf)

            log_output.insert(tk.END, f"Páginas procesada y guardadas en {output_pdf_path}\n")
        except OSError as e:
            log_output.insert(tk.END, f"Error al guardar el archivo {output_pdf_path}: {e}\n")    
            log_output.see(tk.END)


def mover_pdf_a_procesados(pdf_file_path, log_output):

    processed_folder = os.path.join(base_directory, "PROCESADOS")

    if not os.path.exists(processed_folder):
        os.makedirs(processed_folder)
        log_output.insert(tk.END, f"Creando carpeta de procesados: {processed_folder}\n")
        log_output.insert(tk.END, f"Carpeta creada con suceso!\n")

    destino_pdf = os.path.join(
        processed_folder, os.path.basename(pdf_file_path))
    shutil.move(pdf_file_path, destino_pdf)
    log_output.insert(
        tk.END, f"Archivo original movido desde carpeta input_pdf a carpeta PROCESADOS\n")
    log_output.insert(tk.END, f"Rutina de procesamiento finalizada!\n")


class PDFHandler(FileSystemEventHandler):
    def __init__(self, log_output):
        self.log_output = log_output

    def on_created(self, event):
        if event.is_directory:
            return

        if event.src_path.endswith(".pdf"):
            self.log_output.insert(
                tk.END, f"Nuevo archivo PDF detectado: {event.src_path}\n")
            self.log_output.see(tk.END)
            split_pdf_add_img(event.src_path, SIGNATURE_IMAGE, self.log_output)


def start_observer(log_output):

    if not os.path.exists(WATCH_DIRECTORY):
        log_output.insert(tk.END, f"Error: El directorio {WATCH_DIRECTORY} no existe o no es accesible.\n")
        log_output.see(tk.END)
        return

    event_handler = PDFHandler(log_output)
    observer = Observer()
    observer.schedule(event_handler, WATCH_DIRECTORY, recursive=False)

    log_output.insert(
        tk.END, f"Bienvenido al procesador de DSI. Agregar archivo PDF a la carpeta input_pdf\n")
    log_output.see(tk.END)

    try:
        observer.start()
    except Exception as e:
        log_output.insert(tk.END, f"Error al iniciar el observador: {e}\n")
        log_output.see(tk.END)
        return

    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()


def start_gui():
    root = tk.Tk()
    root.title("Monitor de PDFs")

    # Crear un cuadro de texto con scroll para mostrar el log
    log_output = ScrolledText(root, wrap=tk.WORD, width=100, height=30)
    log_output.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    # Iniciar el observador en un hilo separado
    observer_thread = threading.Thread(
        target=start_observer, args=(log_output,))
    observer_thread.daemon = True
    observer_thread.start()

    root.mainloop()


# Iniciar la interfaz gráfica
if __name__ == "__main__":
    start_gui()
