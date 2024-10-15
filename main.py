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

def extract_cuit_pdf(page):
  
  # Leer el PDF
  text = page.extract_text()
  match = re.search(r"CUIT\s*-\s*(\d{2,3}\s*\d{3,4}\s*\d{3,4})", text)
    
  if match:
    cuit = match.group(1)
    cuit_clean = cuit.replace(" ", "")
    return f"{cuit_clean}"
  
  else:
    return None
  

def split_pdf_add_img(pdf_file_path, image_path, log_output):

  reader = PdfReader(pdf_file_path)

  # Leer la firma
  img = Image.open(image_path)
  
  # Obteniendo nombre del archivo pdf original (numero particular)
  pdf_file_name = os.path.basename(pdf_file_path)
  
  for page_num in range(len(reader.pages)):
    
    # Elimina una pagina del PDF original
    page = reader.pages[page_num]
    cuit_code = extract_cuit_pdf(page)
  
    if cuit_code is None:
      log_output.insert(tk.END, "Ultima pagina del archivo - No contiene CUIT\n")
      log_output.see(tk.END)
      continue

    z_directory = "\\\\192.168.0.20\\clientes"
    cuit_folder = os.path.join(z_directory, cuit_code)
  
    if not os.path.exists(cuit_folder):
      os.makedirs(cuit_folder)
      log_output.insert(tk.END, f"Creando carpeta {cuit_folder}\n")
    else:
      log_output.insert(tk.END, f"La carpeta {cuit_folder} ya existe\n")
      
    dsi_folder = os.path.join(cuit_folder, "DSI")
    if not os.path.exists(dsi_folder):
      os.makedirs(dsi_folder)
      log_output.insert(tk.END, f"Creando subcarpeta DSI en {dsi_folder}\n")
      
    time_stamp = time.strftime("%Y%m%d_%H%M%S")
    output_pdf_path = os.path.join(dsi_folder, f"{time_stamp}.pdf")

    #Crea el archivo PDF con la firma en cada pagina
    writer = PdfWriter()

    # Crea un archivo en blanco con las mismas configuraciones de la pagina original
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # Obteniendo tamano de la pagina para centrar la firma
    page_width, page_height = letter
    img_width, img_height = img.size

    # Calculando x e y para centrar la firma
    x = (page_width - img_width) / 3.5
    y = (page_height - img_height) / 3.5

    # Dibujando la firma en el centro de la pagina
    can.drawImage(image_path, x, y, width=img_width, height=img_height)

    can.save()

    # Moviendo el archivo en blanco a la pagina extraida
    packet.seek(0)
    new_pdf = PdfReader(packet)

    page.merge_page(new_pdf.pages[0])
    writer.add_page(page)
    
    with open(output_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)
    
    log_output.insert(tk.END, f"Página {page_num + 1} procesada y guardada en {output_pdf_path}\n")
    log_output.see(tk.END)
    
  log_output.insert(tk.END, f"Procesamiento finalizado\n")
  log_output.see(tk.END)
  
  mover_pdf_a_procesados(pdf_file_path, log_output)

def mover_pdf_a_procesados(pdf_file_path, log_output):
  
  root_directory = os.path.dirname(os.path.dirname(pdf_file_path))
  processed_folder = os.path.join(root_directory, "PROCESADOS")
  
  if not os.path.exists(processed_folder):
    os.makedirs(processed_folder)
    log_output.insert(tk.END, f"Creando carpeta de procesados: {processed_folder}\n")
    log_output.insert(tk.END, f"Carpeta creada con suceso!\n")
    
  destino_pdf = os.path.join(processed_folder, os.path.basename(pdf_file_path))
  shutil.move(pdf_file_path, destino_pdf)
  log_output.insert(tk.END, f"Archivo original movido desde carpeta input_pdf a carpeta PROCESADOS\n")
  log_output.insert(tk.END, f"Rutina de procesamiento finalizada!\n")
  
class PDFHandler(FileSystemEventHandler):
  def __init__(self, log_output):
    self.log_output = log_output
    
  def on_created(self, event):
    if event.is_directory:
      return
    
    if event.src_path.endswith(".pdf"):
      self.log_output.insert(tk.END, f"Nuevo archivo PDF detectado: {event.src_path}\n")
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
  
  log_output.insert(tk.END, f"Bienvenido al procesador de DSI. Agregar archivo PDF a la carpeta input_pdf\n")
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
  observer_thread = threading.Thread(target=start_observer, args=(log_output,))
  observer_thread.daemon = True
  observer_thread.start()
  
  root.mainloop()
  
# Iniciar la interfaz gráfica
if __name__ == "__main__":
    start_gui()

