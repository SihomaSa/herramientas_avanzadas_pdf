import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
from PIL import Image, ImageTk
from datetime import datetime
import io
import pythoncom
import win32com.client  # Para conversión a Word (Windows only)

class PDFToolsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Herramientas PDF Avanzadas")
        self.root.geometry("650x550")
        
        # Variables
        self.pdf_files = []
        self.other_files = []
        self.output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        
          # Fuentes predefinidas
        self.available_fonts = ["helv", "tiro", "cour", "times"]
        self.font_names = {
            "helv": "Helvetica",
            "tiro": "Times-Roman",
            "cour": "Courier",
            "times": "Times-Bold"
        }
        # Crear interfaz
        self.create_widgets()
        
    def create_widgets(self):
        # Barra de menú superior
        menubar = tk.Menu(self.root)
        
        # Menú Archivo
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Salir", command=self.root.quit)
        menubar.add_cascade(label="Archivo", menu=file_menu)
        
        # Menú PDF
        pdf_menu = tk.Menu(menubar, tearoff=0)
        pdf_menu.add_command(label="Unir PDFs", command=self.merge_pdfs)
        pdf_menu.add_command(label="Dividir PDF", command=self.split_pdf)
        pdf_menu.add_command(label="Rotar Páginas", command=self.rotate_pages)
        pdf_menu.add_command(label="Comprimir PDF", command=self.compress_pdf)
        pdf_menu.add_command(label="Numerar Páginas", command=self.number_pages)
        pdf_menu.add_command(label="Editar PDF", command=self.edit_pdf)
        menubar.add_cascade(label="Herramientas PDF", menu=pdf_menu)
        
        # Menú Conversión
        convert_menu = tk.Menu(menubar, tearoff=0)
        convert_menu.add_command(label="PDF a JPG", command=self.pdf_to_jpg)
        convert_menu.add_command(label="JPG a PDF", command=self.jpg_to_pdf)
        convert_menu.add_command(label="JPG a PNG", command=self.jpg_to_png)
        convert_menu.add_command(label="PNG a JPG", command=self.png_to_jpg)
        convert_menu.add_command(label="PDF a Word", command=self.pdf_to_word)
        convert_menu.add_command(label="Word a PDF", command=self.word_to_pdf)
        menubar.add_cascade(label="Conversión", menu=convert_menu)
        
        self.root.config(menu=menubar)
        
        # Frame principal
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        tk.Label(main_frame, text="Herramientas Avanzadas para PDF", font=("Arial", 16, "bold")).pack(pady=10)
        
        # Pestañas
        tab_control = ttk.Notebook(main_frame)
        
        # Pestaña PDF
        pdf_tab = ttk.Frame(tab_control)
        tab_control.add(pdf_tab, text="Operaciones PDF")
        
        # Lista de archivos PDF
        self.pdf_listbox = tk.Listbox(pdf_tab, height=10, selectmode=tk.EXTENDED)
        self.pdf_listbox.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.pdf_listbox, orient="vertical", command=self.pdf_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.pdf_listbox.config(yscrollcommand=scrollbar.set)
        
        # Frame de botones PDF
        pdf_button_frame = tk.Frame(pdf_tab)
        pdf_button_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(pdf_button_frame, text="Agregar PDF(s)", command=self.add_pdfs).pack(side=tk.LEFT, padx=2)
        tk.Button(pdf_button_frame, text="Eliminar", command=lambda: self.remove_files(self.pdf_listbox, self.pdf_files)).pack(side=tk.LEFT, padx=2)
        
        # Frame de botones de operaciones PDF
        pdf_ops_frame = tk.Frame(pdf_tab)
        pdf_ops_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(pdf_ops_frame, text="Unir PDFs", command=self.merge_pdfs, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(pdf_ops_frame, text="Dividir", command=self.split_pdf, bg="#2196F3", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(pdf_ops_frame, text="Rotar", command=self.rotate_pages, bg="#FF9800", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(pdf_ops_frame, text="Comprimir", command=self.compress_pdf, bg="#9C27B0", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(pdf_ops_frame, text="Numerar", command=self.number_pages, bg="#607D8B", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(pdf_ops_frame, text="Editar", command=self.edit_pdf, bg="#795548", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(pdf_ops_frame, text="Limpiar", command=lambda: self.remove_files(self.pdf_listbox, self.pdf_files)).pack(side=tk.LEFT, padx=2)
        tk.Button(pdf_ops_frame, text="Salir", command=self.root.quit, bg="#F44336", fg="white").pack(side=tk.LEFT, padx=2)
        # Pestaña Conversión
        convert_tab = ttk.Frame(tab_control)
        tab_control.add(convert_tab, text="Conversión")
        
        # Lista de archivos para conversión
        self.convert_listbox = tk.Listbox(convert_tab, height=10, selectmode=tk.EXTENDED)
        self.convert_listbox.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        # Scrollbar
        convert_scrollbar = ttk.Scrollbar(self.convert_listbox, orient="vertical", command=self.convert_listbox.yview)
        convert_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.convert_listbox.config(yscrollcommand=convert_scrollbar.set)
        
        # Frame de botones Conversión
        convert_button_frame = tk.Frame(convert_tab)
        convert_button_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(convert_button_frame, text="Agregar Archivo(s)", command=self.add_convert_files).pack(side=tk.LEFT, padx=2)
        tk.Button(convert_button_frame, text="Eliminar", command=lambda: self.remove_files(self.convert_listbox, self.other_files)).pack(side=tk.LEFT, padx=2)
        
        # Frame de botones de conversión
        convert_ops_frame = tk.Frame(convert_tab)
        convert_ops_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(convert_ops_frame, text="PDF a JPG", command=self.pdf_to_jpg, bg="#E91E63", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(convert_ops_frame, text="JPG a PDF", command=self.jpg_to_pdf, bg="#3F51B5", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(convert_ops_frame, text="JPG a PNG", command=self.jpg_to_png, bg="#009688", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(convert_ops_frame, text="PNG a JPG", command=self.png_to_jpg, bg="#FF9800", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(convert_ops_frame, text="PDF a Word", command=self.pdf_to_word, bg="#2907D1", fg="white").pack(side=tk.LEFT, padx=2)
        tk.Button(convert_ops_frame, text="Word a PDF", command=self.word_to_pdf, bg="#FF5722", fg="white").pack(side=tk.LEFT, padx=2)
        
        tab_control.pack(fill=tk.BOTH, expand=True)
        
        # Barra de estado
        self.status_var = tk.StringVar()
        self.status_var.set("Listo. Seleccione archivos y operación")
        tk.Label(main_frame, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X)
    
    def add_pdfs(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos PDF",
            filetypes=(("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*"))
        )
        
        if files:
            for file in files:
                if file not in self.pdf_files:
                    self.pdf_files.append(file)
                    self.pdf_listbox.insert(tk.END, os.path.basename(file))
            
            self.status_var.set(f"{len(self.pdf_files)} PDF(s) listo(s)")
    
    def add_convert_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos",
            filetypes=(("Archivos PDF", "*.pdf"), 
                      ("Archivos JPG", "*.jpg;*.jpeg"),
                      ("Archivos Word", "*.doc;*.docx"),
                      ("Todos los archivos", "*.*"))
        )
        
        if files:
            for file in files:
                if file not in self.other_files:
                    self.other_files.append(file)
                    self.convert_listbox.insert(tk.END, os.path.basename(file))
            
            self.status_var.set(f"{len(self.other_files)} archivo(s) listo(s) para conversión")
    
    def remove_files(self, listbox, file_list):
        selected = listbox.curselection()
        if selected:
            for index in selected[::-1]:  # Eliminar en orden inverso para evitar problemas de índice
                del file_list[index]
                listbox.delete(index)
            
            self.status_var.set(f"{len(file_list)} archivo(s) restante(s)")
        else:
            messagebox.showwarning("Advertencia", "Por favor seleccione archivos para eliminar")
    
    def merge_pdfs(self):
        if len(self.pdf_files) < 2:
            messagebox.showerror("Error", "Debe agregar al menos 2 archivos PDF para unir")
            return
        
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"PDF_Unido_{timestamp}.pdf"
            output_path = os.path.join(self.output_folder, output_filename)
            
            merged_pdf = fitz.open()
            
            for pdf in self.pdf_files:
                doc = fitz.open(pdf)
                merged_pdf.insert_pdf(doc)
                doc.close()
            
            merged_pdf.save(output_path)
            merged_pdf.close()
            
            self.status_var.set(f"PDFs unidos correctamente! Guardado como: {output_filename}")
            messagebox.showinfo("Éxito", f"Archivos PDF unidos correctamente.\nGuardado en: {output_path}")
            
            self.pdf_files.clear()
            self.pdf_listbox.delete(0, tk.END)
            
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al unir los PDFs:\n{str(e)}")
            self.status_var.set("Error al unir PDFs")
    
    def split_pdf(self):
        if not self.pdf_files:
            messagebox.showerror("Error", "Por favor agregue un archivo PDF primero")
            return
        
        if len(self.pdf_files) > 1:
            messagebox.showwarning("Advertencia", "Solo se puede dividir un PDF a la vez. Se usará el primero de la lista")
        
        input_pdf = self.pdf_files[0]
        
        split_window = tk.Toplevel(self.root)
        split_window.title("Dividir PDF")
        split_window.geometry("400x250")
        
        tk.Label(split_window, text=f"Archivo: {os.path.basename(input_pdf)}", font=("Arial", 10)).pack(pady=5)
        
        doc = fitz.open(input_pdf)
        page_count = doc.page_count
        doc.close()
        
        tk.Label(split_window, text=f"El PDF tiene {page_count} páginas").pack()
        tk.Label(split_window, text="Ingrese el rango de páginas a extraer:").pack(pady=5)
        tk.Label(split_window, text="Ejemplo: 1-3,5,7-9 (para páginas 1,2,3,5,7,8,9)").pack()
        
        page_range_entry = tk.Entry(split_window, width=30)
        page_range_entry.pack(pady=10)
        page_range_entry.focus_set()
        
        def perform_split():
            page_range = page_range_entry.get()
            if not page_range:
                messagebox.showerror("Error", "Debe ingresar un rango de páginas")
                return
            
            try:
                doc = fitz.open(input_pdf)
                output_filename = f"Dividido_{os.path.basename(input_pdf)}"
                output_path = os.path.join(self.output_folder, output_filename)
                
                pages = self.parse_page_range(page_range, page_count)
                
                new_doc = fitz.open()
                for page_num in pages:
                    new_doc.insert_pdf(doc, from_page=page_num-1, to_page=page_num-1)
                
                new_doc.save(output_path)
                new_doc.close()
                doc.close()
                
                messagebox.showinfo("Éxito", f"PDF dividido correctamente.\nGuardado en: {output_path}")
                split_window.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al dividir PDF:\n{str(e)}")
        
        tk.Button(split_window, text="Dividir", command=perform_split, bg="#2196F3", fg="white").pack(pady=10)
        tk.Button(split_window, text="Cancelar", command=split_window.destroy).pack()
    
    def rotate_pages(self):
        if not self.pdf_files:
            messagebox.showerror("Error", "Por favor agregue un archivo PDF primero")
            return
        
        if len(self.pdf_files) > 1:
            messagebox.showwarning("Advertencia", "Solo se puede rotar un PDF a la vez. Se usará el primero de la lista")
        
        input_pdf = self.pdf_files[0]
        
        rotate_window = tk.Toplevel(self.root)
        rotate_window.title("Rotar PDF")
        rotate_window.geometry("400x300")
        
        tk.Label(rotate_window, text=f"Archivo: {os.path.basename(input_pdf)}", font=("Arial", 10)).pack(pady=5)
        
        doc = fitz.open(input_pdf)
        page_count = doc.page_count
        doc.close()
        
        tk.Label(rotate_window, text=f"El PDF tiene {page_count} páginas").pack()
        
        tk.Label(rotate_window, text="Ángulo de rotación:").pack(pady=5)
        
        angle_var = tk.IntVar(value=90)
        angle_frame = tk.Frame(rotate_window)
        angle_frame.pack()
        
        tk.Radiobutton(angle_frame, text="90°", variable=angle_var, value=90).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(angle_frame, text="180°", variable=angle_var, value=180).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(angle_frame, text="270°", variable=angle_var, value=270).pack(side=tk.LEFT, padx=5)
        
        tk.Label(rotate_window, text="Páginas a rotar (opcional):").pack(pady=5)
        tk.Label(rotate_window, text="Ejemplo: 1-3,5,7-9 (dejar vacío para todas)").pack()
        
        page_range_entry = tk.Entry(rotate_window, width=30)
        page_range_entry.pack(pady=10)
        
        def perform_rotation():
            angle = angle_var.get()
            page_range = page_range_entry.get()
            
            try:
                doc = fitz.open(input_pdf)
                output_filename = f"Rotado_{os.path.basename(input_pdf)}"
                output_path = os.path.join(self.output_folder, output_filename)
                
                if page_range:
                    pages = self.parse_page_range(page_range, page_count)
                else:
                    pages = range(1, page_count + 1)
                
                for page_num in pages:
                    page = doc[page_num-1]
                    page.set_rotation(angle)
                
                doc.save(output_path)
                doc.close()
                
                messagebox.showinfo("Éxito", f"PDF rotado correctamente.\nGuardado en: {output_path}")
                rotate_window.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al rotar PDF:\n{str(e)}")
        
        tk.Button(rotate_window, text="Rotar", command=perform_rotation, bg="#FF9800", fg="white").pack(pady=5)
        tk.Button(rotate_window, text="Cancelar", command=rotate_window.destroy).pack()
    
    def compress_pdf(self):
        if not self.pdf_files:
            messagebox.showerror("Error", "Por favor agregue un archivo PDF primero")
            return
        
        if len(self.pdf_files) > 1:
            messagebox.showwarning("Advertencia", "Solo se puede comprimir un PDF a la vez. Se usará el primero de la lista")
        
        input_pdf = self.pdf_files[0]
        original_size = os.path.getsize(input_pdf) / (1024 * 1024)  # Tamaño en MB
        
        compress_window = tk.Toplevel(self.root)
        compress_window.title("Comprimir PDF")
        compress_window.geometry("450x300")
        
        tk.Label(compress_window, text=f"Archivo: {os.path.basename(input_pdf)}", font=("Arial", 10)).pack(pady=5)
        tk.Label(compress_window, text=f"Tamaño actual: {original_size:.2f} MB").pack()
        
        tk.Label(compress_window, text="Seleccione el nivel de compresión:").pack(pady=10)
        
        level_var = tk.StringVar(value="medio")
        
        levels_frame = tk.Frame(compress_window)
        levels_frame.pack()
        
        tk.Radiobutton(levels_frame, text="Alta compresión", variable=level_var, value="alta").pack(anchor=tk.W)
        tk.Label(levels_frame, text="(Tamaño más pequeño, posible pérdida de calidad)").pack(anchor=tk.W, padx=20)
        
        tk.Radiobutton(levels_frame, text="Compresión media", variable=level_var, value="medio").pack(anchor=tk.W, pady=5)
        tk.Label(levels_frame, text="(Balance entre tamaño y calidad)").pack(anchor=tk.W, padx=20)
        
        tk.Radiobutton(levels_frame, text="Baja compresión", variable=level_var, value="baja").pack(anchor=tk.W)
        tk.Label(levels_frame, text="(Mínima compresión, máxima calidad)").pack(anchor=tk.W, padx=20)
        
        def perform_compression():
            level = level_var.get()
            
            try:
                doc = fitz.open(input_pdf)
                output_filename = f"Comprimido_{os.path.basename(input_pdf)}"
                output_path = os.path.join(self.output_folder, output_filename)
                
                if level == "alta":
                    compress = True
                    garbage = 4
                    clean = True
                elif level == "medio":
                    compress = True
                    garbage = 2
                    clean = True
                else:  # baja
                    compress = True
                    garbage = 0
                    clean = False
                
                doc.save(output_path, garbage=garbage, clean=clean, deflate=compress)
                doc.close()
                
                new_size = os.path.getsize(output_path) / (1024 * 1024)  # MB
                reduction = (original_size - new_size) / original_size * 100
                
                messagebox.showinfo("Éxito", 
                    f"PDF comprimido correctamente.\n\n"
                    f"Tamaño original: {original_size:.2f} MB\n"
                    f"Nuevo tamaño: {new_size:.2f} MB\n"
                    f"Reducción: {reduction:.1f}%\n\n"
                    f"Guardado en: {output_path}")
                
                compress_window.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al comprimir PDF:\n{str(e)}")
        
        tk.Button(compress_window, text="Comprimir", command=perform_compression, bg="#9C27B0", fg="white").pack(pady=10)
        tk.Button(compress_window, text="Cancelar", command=compress_window.destroy).pack()
    
    def pdf_to_jpg(self):
        if not self.other_files:
            messagebox.showerror("Error", "Por favor agregue al menos un archivo PDF")
            return
        
        # Crear ventana de configuración
        config_window = tk.Toplevel(self.root)
        config_window.title("PDF a JPG - Configuración")
        config_window.geometry("400x250")
        
        tk.Label(config_window, text="Configuración de conversión PDF a JPG", font=("Arial", 10, "bold")).pack(pady=5)
        
        # Calidad de imagen
        tk.Label(config_window, text="Calidad de imagen (1-100):").pack(pady=5)
        quality_var = tk.IntVar(value=90)
        tk.Scale(config_window, from_=1, to=100, orient=tk.HORIZONTAL, variable=quality_var).pack()
        
        # DPI
        tk.Label(config_window, text="Resolución (DPI):").pack(pady=5)
        dpi_var = tk.IntVar(value=300)
        tk.Entry(config_window, textvariable=dpi_var, width=10).pack()
        
        def perform_conversion():
            quality = quality_var.get()
            dpi = dpi_var.get()
            
            if quality < 1 or quality > 100:
                messagebox.showerror("Error", "La calidad debe estar entre 1 y 100")
                return
            
            if dpi < 72 or dpi > 1200:
                messagebox.showerror("Error", "El DPI debe estar entre 72 y 1200")
                return
            
            try:
                for pdf_path in self.other_files:
                    if not pdf_path.lower().endswith('.pdf'):
                        continue
                    
                    # Crear carpeta para las imágenes
                    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                    output_folder = os.path.join(self.output_folder, f"PDF_a_JPG_{base_name}")
                    os.makedirs(output_folder, exist_ok=True)
                    
                    # Abrir PDF
                    pdf_document = fitz.open(pdf_path)
                    
                    for page_num in range(len(pdf_document)):
                        page = pdf_document.load_page(page_num)
                        zoom = dpi / 72  # Ajustar según DPI
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat)
                        
                        # Convertir a imagen PIL
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        
                        # Guardar JPG
                        output_path = os.path.join(output_folder, f"{base_name}_pagina_{page_num+1}.jpg")
                        img.save(output_path, "JPEG", quality=quality)
                    
                    pdf_document.close()
                
                messagebox.showinfo("Éxito", f"Conversión completada.\nLas imágenes se guardaron en:\n{output_folder}")
                config_window.destroy()
                self.other_files.clear()
                self.convert_listbox.delete(0, tk.END)
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al convertir PDF a JPG:\n{str(e)}")
        
        tk.Button(config_window, text="Convertir", command=perform_conversion, bg="#E91E63", fg="white").pack(pady=10)
        tk.Button(config_window, text="Cancelar", command=config_window.destroy).pack()
    
    def jpg_to_pdf(self):
        if not self.other_files:
            messagebox.showerror("Error", "Por favor agregue al menos un archivo JPG")
            return
        
        # Filtrar solo archivos de imagen
        image_files = [f for f in self.other_files if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
        
        if not image_files:
            messagebox.showerror("Error", "No se encontraron archivos JPG/JPEG para convertir")
            return
        
        try:
            # Crear un nuevo PDF
            pdf = fitz.open()
            
            for img_path in image_files:
                # Abrir imagen con PIL
                img = Image.open(img_path)
                
                # Convertir a PDF page
                pdf_bytes = io.BytesIO()
                img.save(pdf_bytes, "PDF", resolution=100.0)
                pdf_bytes.seek(0)
                
                # Insertar en el PDF
                img_pdf = fitz.open("pdf", pdf_bytes.read())
                pdf.insert_pdf(img_pdf)
                img_pdf.close()
            
            # Guardar PDF
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"JPG_a_PDF_{timestamp}.pdf"
            output_path = os.path.join(self.output_folder, output_filename)
            
            pdf.save(output_path)
            pdf.close()
            
            messagebox.showinfo("Éxito", f"Conversión completada.\nPDF guardado en:\n{output_path}")
            self.other_files.clear()
            self.convert_listbox.delete(0, tk.END)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al convertir JPG a PDF:\n{str(e)}")
    # Convertir JPG a PNG
    def jpg_to_png(self):
        if not self.other_files:
            messagebox.showerror("Error", "Por favor agregue al menos un archivo JPG")
            return
        
        # Filtrar solo archivos JPG
        jpg_files = [f for f in self.other_files if f.lower().endswith(('.jpg', '.jpeg'))]
        
        if not jpg_files:
            messagebox.showerror("Error", "No se encontraron archivos JPG para convertir")
            return
        
        try:
            for jpg_path in jpg_files:
                # Abrir imagen JPG
                img = Image.open(jpg_path)
                
                # Crear nombre de archivo de salida
                base_name = os.path.splitext(os.path.basename(jpg_path))[0]
                output_filename = f"{base_name}.png"
                output_path = os.path.join(self.output_folder, output_filename)
                
                # Guardar como PNG
                img.save(output_path, "PNG")
            
            messagebox.showinfo("Éxito", f"Conversión completada.\nArchivos PNG guardados en:\n{self.output_folder}")
            self.other_files.clear()
            self.convert_listbox.delete(0, tk.END)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al convertir JPG a PNG:\n{str(e)}")
    # Convertir PNG a JPG    
    def png_to_jpg(self):
        if not self.other_files:
            messagebox.showerror("Error", "Por favor agregue al menos un archivo PNG")
            return
        
        # Filtrar solo archivos PNG
        png_files = [f for f in self.other_files if f.lower().endswith('.png')]
        
        if not png_files:
            messagebox.showerror("Error", "No se encontraron archivos PNG para convertir")
            return
        
        # Ventana de configuración para calidad JPG
        config_window = tk.Toplevel(self.root)
        config_window.title("Configuración PNG a JPG")
        config_window.geometry("300x150")
        
        tk.Label(config_window, text="Calidad de imagen (1-100):").pack(pady=5)
        quality_var = tk.IntVar(value=90)
        tk.Scale(config_window, from_=1, to=100, orient=tk.HORIZONTAL, variable=quality_var).pack()
        
        def perform_conversion():
            quality = quality_var.get()
            
            try:
                for png_path in png_files:
                    # Abrir imagen PNG
                    img = Image.open(png_path)
                    
                    # Convertir a RGB si tiene transparencia
                    if img.mode in ('RGBA', 'LA'):
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        background.paste(img, mask=img.split()[-1])
                        img = background
                    
                    # Crear nombre de archivo de salida
                    base_name = os.path.splitext(os.path.basename(png_path))[0]
                    output_filename = f"{base_name}.jpg"
                    output_path = os.path.join(self.output_folder, output_filename)
                    
                    # Guardar como JPG
                    img.save(output_path, "JPEG", quality=quality)
                
                messagebox.showinfo("Éxito", f"Conversión completada.\nArchivos JPG guardados en:\n{self.output_folder}")
                config_window.destroy()
                self.other_files.clear()
                self.convert_listbox.delete(0, tk.END)
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al convertir PNG a JPG:\n{str(e)}")
        
        tk.Button(config_window, text="Convertir", command=perform_conversion, bg="#4CAF50", fg="white").pack(pady=10)
        tk.Button(config_window, text="Cancelar", command=config_window.destroy).pack() 
        
    
    def pdf_to_word(self):
        if not self.other_files:
            messagebox.showerror("Error", "Por favor agregue al menos un archivo PDF")
            return
        
        # Filtrar solo archivos PDF
        pdf_files = [f for f in self.other_files if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            messagebox.showerror("Error", "No se encontraron archivos PDF para convertir")
            return
        
        try:
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            for pdf_path in pdf_files:
                # Crear nombre de archivo de salida
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                output_filename = f"{base_name}.docx"
                output_path = os.path.join(self.output_folder, output_filename)
                
                # Convertir PDF a Word
                doc = word.Documents.Open(pdf_path)
                doc.SaveAs(output_path, FileFormat=16)  # 16 = wdFormatDocumentDefault
                doc.Close()
            
            word.Quit()
            pythoncom.CoUninitialize()
            
            messagebox.showinfo("Éxito", f"Conversión completada.\nDocumentos Word guardados en:\n{self.output_folder}")
            self.other_files.clear()
            self.convert_listbox.delete(0, tk.END)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al convertir PDF a Word:\n{str(e)}")
            try:
                word.Quit()
                pythoncom.CoUninitialize()
            except:
                pass
    
    def word_to_pdf(self):
        if not self.other_files:
            messagebox.showerror("Error", "Por favor agregue al menos un archivo Word")
            return
        
        # Filtrar solo archivos Word
        word_files = [f for f in self.other_files if f.lower().endswith(('.doc', '.docx'))]
        
        if not word_files:
            messagebox.showerror("Error", "No se encontraron archivos Word para convertir")
            return
        
        try:
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            for word_path in word_files:
                # Crear nombre de archivo de salida
                base_name = os.path.splitext(os.path.basename(word_path))[0]
                output_filename = f"{base_name}.pdf"
                output_path = os.path.join(self.output_folder, output_filename)
                
                # Convertir Word a PDF
                doc = word.Documents.Open(word_path)
                doc.SaveAs(output_path, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
            
            word.Quit()
            pythoncom.CoUninitialize()
            
            messagebox.showinfo("Éxito", f"Conversión completada.\nPDFs guardados en:\n{self.output_folder}")
            self.other_files.clear()
            self.convert_listbox.delete(0, tk.END)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al convertir Word a PDF:\n{str(e)}")
            try:
                word.Quit()
                pythoncom.CoUninitialize()
            except:
                pass

    def number_pages(self):
            if not self.pdf_files:
                messagebox.showerror("Error", "Por favor agregue un archivo PDF primero")
                return
            
            if len(self.pdf_files) > 1:
                messagebox.showwarning("Advertencia", "Solo se puede numerar un PDF a la vez. Se usará el primero de la lista")
            
            input_pdf = self.pdf_files[0]
            
            number_window = tk.Toplevel(self.root)
            number_window.title("Numerar Páginas")
            number_window.geometry("500x450")
            
            tk.Label(number_window, text=f"Archivo: {os.path.basename(input_pdf)}", font=("Arial", 10, "bold")).pack(pady=5)
            
            # Configuración de numeración
            tk.Label(number_window, text="Formato del número de página:").pack(anchor=tk.W, padx=10)
            
            format_var = tk.StringVar(value="{page} de {total}")
            formats = [
                "{page}",
                "{page} de {total}",
                "Página {page}",
                "- {page} -",
                "Sección {section} - {page}"
            ]
            
            format_frame = tk.Frame(number_window)
            format_frame.pack(fill=tk.X, padx=10, pady=5)
            
            for fmt in formats:
                tk.Radiobutton(format_frame, text=fmt, variable=format_var, value=fmt).pack(anchor=tk.W)
            
            # Posición del número
            tk.Label(number_window, text="Posición en la página:").pack(anchor=tk.W, padx=10, pady=(10,0))
            
            position_var = tk.StringVar(value="bottom-center")
            positions = [
                ("Abajo-Centro", "bottom-center"),
                ("Abajo-Derecha", "bottom-right"),
                ("Abajo-Izquierda", "bottom-left"),
                ("Arriba-Centro", "top-center"),
                ("Arriba-Derecha", "top-right"),
                ("Arriba-Izquierda", "top-left")
            ]
            
            position_frame = tk.Frame(number_window)
            position_frame.pack(fill=tk.X, padx=10, pady=5)
            
            for text, pos in positions:
                tk.Radiobutton(position_frame, text=text, variable=position_var, value=pos).pack(side=tk.LEFT, padx=5)
            
            # Configuración de fuente y color primero
            tk.Label(number_window, text="Apariencia del número:").pack(anchor=tk.W, padx=10, pady=(10,0))
            
            appearance_frame = tk.Frame(number_window)
            appearance_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # Fuente
            tk.Label(appearance_frame, text="Fuente:").pack(side=tk.LEFT, padx=5)
            font_var = tk.StringVar(value="helv")
            font_menu = tk.OptionMenu(appearance_frame, font_var, *self.available_fonts)
            font_menu.pack(side=tk.LEFT, padx=5)
            
            # Tamaño
            tk.Label(appearance_frame, text="Tamaño:").pack(side=tk.LEFT, padx=5)
            size_var = tk.IntVar(value=12)
            tk.Spinbox(appearance_frame, from_=8, to=72, width=3, textvariable=size_var).pack(side=tk.LEFT, padx=5)
            
            # Color
            tk.Label(appearance_frame, text="Color:").pack(side=tk.LEFT, padx=5)
            color_var = tk.StringVar(value="#000000")
            color_entry = tk.Entry(appearance_frame, textvariable=color_var, width=7)
            color_entry.pack(side=tk.LEFT, padx=5)
            tk.Button(appearance_frame, text="Seleccionar", command=lambda: self.choose_color(color_var)).pack(side=tk.LEFT, padx=5)
            
            # Margen
            tk.Label(number_window, text="Margen desde el borde (puntos):").pack(anchor=tk.W, padx=10, pady=(10,0))
            margin_var = tk.IntVar(value=20)
            tk.Scale(number_window, from_=0, to=100, orient=tk.HORIZONTAL, variable=margin_var).pack(fill=tk.X, padx=10)
        
    def perform_numbering():
            fmt = format_var.get()
            position = position_var.get()
            font = font_var.get()
            size = size_var.get()
            color = color_var.get()
            margin = margin_var.get()
            
            try:
                # Convertir color hex a RGB
                if color.startswith("#"):
                    color = color.lstrip("#")
                    color = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
                else:
                    color = (0, 0, 0)  # Negro por defecto
                
                doc = fitz.open(input_pdf)
                total_pages = len(doc)
                
                for page_num in range(total_pages):
                    page = doc[page_num]
                    width = page.rect.width
                    height = page.rect.height
                    
                    # Determinar posición
                    if position == "bottom-center":
                        x = width / 2
                        y = height - margin
                        align = 1  # Centro
                    elif position == "bottom-right":
                        x = width - margin
                        y = height - margin
                        align = 2  # Derecha
                    elif position == "bottom-left":
                        x = margin
                        y = height - margin
                        align = 0  # Izquierda
                    elif position == "top-center":
                        x = width / 2
                        y = margin
                        align = 1
                    elif position == "top-right":
                        x = width - margin
                        y = margin
                        align = 2
                    elif position == "top-left":
                        x = margin
                        y = margin
                        align = 0
                    
                    # Formatear texto
                    text = fmt.replace("{page}", str(page_num + 1))
                    text = text.replace("{total}", str(total_pages))
                    text = text.replace("{section}", str((page_num // 10) + 1))
                    
                    # Insertar número de página
                    page.insert_text(
                        point=(x, y),
                        text=text,
                        fontname=font,
                        fontsize=size,
                        color=color,
                        align=align
                    )
                
                # Guardar el PDF con números
                output_filename = f"Numerado_{os.path.basename(input_pdf)}"
                output_path = os.path.join(self.output_folder, output_filename)
                doc.save(output_path)
                doc.close()
                
                messagebox.showinfo("Éxito", f"Páginas numeradas correctamente.\nGuardado en: {output_path}")
                number_window.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al numerar páginas:\n{str(e)}")
        
            tk.Button(number_window, text="Aplicar Numeración", command=perform_numbering, bg="#607D8B", fg="white").pack(pady=10)
            tk.Button(number_window, text="Cancelar", command=number_window.destroy).pack()
    
    def edit_pdf(self):
        if not self.pdf_files:
            messagebox.showerror("Error", "Por favor agregue un archivo PDF primero")
            return
        
        if len(self.pdf_files) > 1:
            messagebox.showwarning("Advertencia", "Solo se puede editar un PDF a la vez. Se usará el primero de la lista")
        
        input_pdf = self.pdf_files[0]
        self.current_pdf = input_pdf
        self.current_page = 0
        self.doc = fitz.open(input_pdf)
        self.total_pages = len(self.doc)
        self.zoom_level = 1.0  # Nivel de zoom inicial
        self.text_annotations = []  # Almacenar las anotaciones de texto
        
        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"Editor de PDF - {os.path.basename(input_pdf)}")
        edit_window.geometry("1000x700")
        
        # Frame principal con paneles divididos
        main_frame = tk.Frame(edit_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Panel izquierdo para herramientas (30% del ancho)
        tools_frame = tk.Frame(main_frame, width=300, bg="#f0f0f0")
        tools_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        tools_frame.pack_propagate(False)
        
        # Panel derecho para visualización del PDF (70% del ancho)
        pdf_frame = tk.Frame(main_frame)
        pdf_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Controles de navegación en el panel de herramientas
        nav_frame = tk.Frame(tools_frame)
        nav_frame.pack(fill=tk.X, pady=10)
        
        tk.Button(nav_frame, text="← Anterior", 
                command=lambda: self.change_page(-1, edit_window)).pack(side=tk.LEFT, padx=2)
        self.page_label = tk.Label(nav_frame, text=f"Página 1 de {self.total_pages}")
        self.page_label.pack(side=tk.LEFT, padx=10)
        tk.Button(nav_frame, text="Siguiente →", 
                command=lambda: self.change_page(1, edit_window)).pack(side=tk.LEFT, padx=2)
        
        # Controles de zoom en el panel de herramientas
        zoom_frame = tk.Frame(tools_frame)
        zoom_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(zoom_frame, text="Zoom:").pack(side=tk.LEFT, padx=5)
        tk.Button(zoom_frame, text="-", command=lambda: self.adjust_zoom(-0.1, edit_window)).pack(side=tk.LEFT, padx=2)
        tk.Button(zoom_frame, text="+", command=lambda: self.adjust_zoom(0.1, edit_window)).pack(side=tk.LEFT, padx=2)
        self.zoom_label = tk.Label(zoom_frame, text="100%")
        self.zoom_label.pack(side=tk.LEFT, padx=5)
        
        # Herramientas de edición de texto en el panel de herramientas
        text_frame = tk.LabelFrame(tools_frame, text="Añadir Texto", padx=5, pady=5)
        text_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(text_frame, text="Texto:").pack(anchor=tk.W)
        self.text_entry = tk.Entry(text_frame)
        self.text_entry.pack(fill=tk.X, pady=2)
        
        font_frame = tk.Frame(text_frame)
        font_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(font_frame, text="Fuente:").pack(side=tk.LEFT)
        self.font_var = tk.StringVar(value="helv")
        tk.OptionMenu(font_frame, self.font_var, *self.available_fonts).pack(side=tk.LEFT)
        
        tk.Label(font_frame, text="Tamaño:").pack(side=tk.LEFT, padx=(10,0))
        self.size_var = tk.IntVar(value=12)
        tk.Spinbox(font_frame, from_=8, to=72, width=3, textvariable=self.size_var).pack(side=tk.LEFT)
        
        color_frame = tk.Frame(text_frame)
        color_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(color_frame, text="Color:").pack(side=tk.LEFT)
        self.color_var = tk.StringVar(value="#000000")
        tk.Entry(color_frame, textvariable=self.color_var, width=7).pack(side=tk.LEFT, padx=5)
        tk.Button(color_frame, text="Seleccionar", 
                command=lambda: self.choose_color(self.color_var)).pack(side=tk.LEFT)
        
        tk.Button(text_frame, text="Añadir Texto al PDF", 
                command=lambda: self.add_text_to_pdf(edit_window), bg="#4CAF50", fg="white").pack(fill=tk.X, pady=5)
        
        # Botones de acción en el panel de herramientas
        action_frame = tk.Frame(tools_frame)
        action_frame.pack(fill=tk.X, pady=10)
        
        tk.Button(action_frame, text="Guardar Cambios", 
                command=lambda: self.save_pdf(edit_window), bg="#2196F3", fg="white").pack(fill=tk.X, pady=2)
        tk.Button(action_frame, text="Guardar Como", 
                command=lambda: self.save_pdf_as(edit_window), bg="#FF9800", fg="white").pack(fill=tk.X, pady=2)
        tk.Button(action_frame, text="Cancelar", 
                command=edit_window.destroy).pack(fill=tk.X, pady=2)
        
        # Visualización del PDF con scrollbars
        pdf_container = tk.Frame(pdf_frame)
        pdf_container.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        h_scroll = tk.Scrollbar(pdf_container, orient=tk.HORIZONTAL)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        v_scroll = tk.Scrollbar(pdf_container)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Canvas para el PDF
        self.pdf_canvas = tk.Canvas(
            pdf_container,
            bg='white',
            xscrollcommand=h_scroll.set,
            yscrollcommand=v_scroll.set
        )
        self.pdf_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configurar scrollbars
        h_scroll.config(command=self.pdf_canvas.xview)
        v_scroll.config(command=self.pdf_canvas.yview)
        
        # Mostrar la primera página
        self.show_page(edit_window)

    def adjust_zoom(self, delta, window):
        """Ajusta el nivel de zoom"""
        self.zoom_level = max(0.1, min(3.0, self.zoom_level + delta))  # Limitar entre 10% y 300%
        self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
        self.show_page(window)

    def show_page(self, window):
        """Muestra la página actual del PDF en el canvas"""
        # Limpiar el canvas
        self.pdf_canvas.delete("all")
        
        # Obtener la página actual
        page = self.doc.load_page(self.current_page)
        zoom = 1.5 * self.zoom_level  # Factor de zoom con ajuste
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # Convertir a formato que tkinter pueda mostrar
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.tk_img = ImageTk.PhotoImage(image=img)
        
        # Mostrar la imagen en el canvas
        self.pdf_canvas.create_image(20, 20, anchor=tk.NW, image=self.tk_img)
        
        # Mostrar las anotaciones de texto para esta página
        for annotation in self.text_annotations:
            if annotation['page'] == self.current_page:
                x, y = annotation['position']
                self.pdf_canvas.create_text(
                    x * zoom + 20, y * zoom + 20,
                    text=annotation['text'],
                    fill=annotation['color'],
                    font=(annotation['font'], annotation['size'])
                )
        
        self.pdf_canvas.config(scrollregion=self.pdf_canvas.bbox(tk.ALL))
        
        # Actualizar etiqueta de página
        self.page_label.config(text=f"Página {self.current_page + 1} de {self.total_pages}")
        
        # Ajustar tamaño de la ventana si es necesario
        window.geometry(f"{min(1000, pix.width + 40)}x{min(800, pix.height + 200)}")

    def change_page(self, delta, window):
        """Cambia a la página anterior o siguiente"""
        new_page = self.current_page + delta
        if 0 <= new_page < self.total_pages:
            self.current_page = new_page
            self.show_page(window)

    def add_text_to_pdf(self, window):
        text = self.text_entry.get()
        if not text:
            messagebox.showwarning("Advertencia", "Por favor ingrese un texto")
            return
        
        try:
            # Obtener posición del clic en el canvas
            def on_canvas_click(event):
                # Convertir coordenadas del canvas a coordenadas del PDF
                x = (event.x - 20) / (1.5 * self.zoom_level)
                y = (event.y - 20) / (1.5 * self.zoom_level)
                
                # Guardar la anotación
                annotation = {
                    'page': self.current_page,
                    'text': text,
                    'position': (x, y),
                    'font': self.font_var.get(),
                    'size': self.size_var.get(),
                    'color': self.color_var.get()
                }
                self.text_annotations.append(annotation)
                
                # Mostrar el texto inmediatamente
                self.show_page(window)
                
                # Remover el binding del clic
                self.pdf_canvas.unbind("<Button-1>")
                messagebox.showinfo("Información", f"Texto añadido a la página {self.current_page + 1}")
            
            # Configurar el binding para el clic
            self.pdf_canvas.bind("<Button-1>", on_canvas_click)
            messagebox.showinfo("Instrucción", "Haga clic en la posición donde desea añadir el texto")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al añadir texto:\n{str(e)}")

    def save_pdf(self, window):
        try:
            # Aplicar los cambios al PDF
            for annotation in self.text_annotations:
                page = self.doc.load_page(annotation['page'])
                rect = fitz.Rect(annotation['position'][0], annotation['position'][1], 
                            annotation['position'][0] + 100, annotation['position'][1] + 20)
                
                # Crear un widget de texto
                annot = page.add_freetext_annot(rect, annotation['text'], 
                                            fontsize=annotation['size'],
                                            fontname=annotation['font'],
                                            text_color=annotation['color'],
                                            fill_color=None)
                annot.update()
            
            # Guardar el PDF
            self.doc.save(self.current_pdf)
            messagebox.showinfo("Éxito", "Cambios guardados correctamente")
            window.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar PDF:\n{str(e)}")

    def save_pdf_as(self, window):
        output_path = filedialog.asksaveasfilename(
            title="Guardar PDF como",
            defaultextension=".pdf",
            filetypes=(("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*"))
        )
        
        if output_path:
            try:
                # Aplicar los cambios al PDF
                for annotation in self.text_annotations:
                    page = self.doc.load_page(annotation['page'])
                    rect = fitz.Rect(annotation['position'][0], annotation['position'][1], 
                                annotation['position'][0] + 100, annotation['position'][1] + 20)
                    
                    # Crear un widget de texto
                    annot = page.add_freetext_annot(rect, annotation['text'], 
                                                fontsize=annotation['size'],
                                                fontname=annotation['font'],
                                                text_color=annotation['color'],
                                                fill_color=None)
                    annot.update()
                
                # Guardar el PDF
                self.doc.save(output_path)
                messagebox.showinfo("Éxito", f"PDF guardado como:\n{output_path}")
                window.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar PDF:\n{str(e)}")

    def choose_color(self, color_var):
        color = tk.colorchooser.askcolor(title="Seleccione un color")
        if color[1]:  # Si se seleccionó un color
            color_var.set(color[1])
    def parse_page_range(self, range_str, max_pages):
        """Convierte un string de rango de páginas en una lista de números de página"""
        pages = []
        parts = range_str.split(',')
        
        for part in parts:
            if '-' in part:
                start, end = map(int, part.split('-'))
                pages.extend(range(start, end+1))
            else:
                pages.append(int(part))
        
        # Validar páginas
        for page in pages:
            if page < 1 or page > max_pages:
                raise ValueError(f"Página {page} está fuera de rango (1-{max_pages})")
        
        return sorted(list(set(pages)))  # Eliminar duplicados y ordenar

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolsApp(root)
    root.mainloop()