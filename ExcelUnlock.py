import os
import zipfile
import re
import xml.etree.ElementTree as ET
from tempfile import TemporaryDirectory
import tkinter as tk
from tkinter import filedialog, messagebox

class ExcelUnlockerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Sheet Unlocker - Offline")
        self.root.geometry("500x520") # Aumentado ligeramente para el autor
        
        self.ruta_archivo = tk.StringVar()
        self.last_dir = os.path.expanduser("~")
        self.hojas_disponibles = [] 
        
        self.setup_ui()

    def setup_ui(self):
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)

        # Selección de archivo
        tk.Label(main_frame, text="Archivo Excel (.xlsx):", font=("Arial", 10, "bold")).pack(anchor="w")
        
        frame_search = tk.Frame(main_frame)
        frame_search.pack(fill="x", pady=(5, 15))
        
        self.entry_ruta = tk.Entry(frame_search, textvariable=self.ruta_archivo, state="readonly")
        self.entry_ruta.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        btn_browse = tk.Button(frame_search, text="Explorar...", command=self.seleccionar_archivo)
        btn_browse.pack(side="right")

        # Lista de hojas
        tk.Label(main_frame, text="Hojas en el archivo (puedes elegir varias):", font=("Arial", 10, "bold")).pack(anchor="w")
        
        list_frame = tk.Frame(main_frame)
        list_frame.pack(fill="both", expand=True, pady=5)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.lista_hojas = tk.Listbox(list_frame, font=("Arial", 10), selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set)
        self.lista_hojas.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.lista_hojas.yview)

        # Botón de acción
        self.btn_desbloquear = tk.Button(main_frame, text="DESBLOQUEAR HOJAS SELECCIONADAS", 
                                       command=self.procesar_desbloqueo, 
                                       state="disabled", 
                                       bg="#2ecc71", fg="white", 
                                       font=("Arial", 10, "bold"),
                                       pady=10)
        self.btn_desbloquear.pack(fill="x", pady=(15, 10))

        # Autor
        lbl_autor = tk.Label(main_frame, text="Autor: Cyberdemon19", font=("Arial", 9, "italic"), fg="gray")
        lbl_autor.pack(pady=5)

    def seleccionar_archivo(self):
        file_path = filedialog.askopenfilename(
            initialdir=self.last_dir,
            title="Seleccionar archivo Excel",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if file_path:
            self.last_dir = os.path.dirname(file_path)
            self.ruta_archivo.set(file_path)
            self.cargar_hojas(file_path)
        
        self.root.update_idletasks()

    def cargar_hojas(self, file_path):
        self.lista_hojas.delete(0, tk.END)
        self.hojas_disponibles = []
        
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
                
                workbook_data = zip_ref.read('xl/workbook.xml')
                root_wb = ET.fromstring(workbook_data)
                
                rels_data = zip_ref.read('xl/_rels/workbook.xml.rels')
                root_rels = ET.fromstring(rels_data)
                rel_map = {rel.get('Id'): rel.get('Target') for rel in root_rels.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')}

                for sheet in root_wb.findall('.//main:sheet', ns):
                    nombre = sheet.get('name')
                    rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    ruta = rel_map.get(rid)
                    if ruta and not ruta.startswith('xl/'): ruta = 'xl/' + ruta
                    
                    self.hojas_disponibles.append((nombre, ruta))
                    self.lista_hojas.insert(tk.END, f"  📄 {nombre}")
            
            self.btn_desbloquear.config(state="normal")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")

    def procesar_desbloqueo(self):
        seleccion = self.lista_hojas.curselection()
        if not seleccion:
            messagebox.showwarning("Selección", "Por favor, selecciona al menos una hoja.")
            return
        
        archivo_path = self.ruta_archivo.get()
        nombre_base = os.path.splitext(archivo_path)[0]
        hojas_ok = []
        
        try:
            with TemporaryDirectory() as tmpdir:
                with zipfile.ZipFile(archivo_path, 'r') as zip_ref:
                    zip_ref.extractall(tmpdir)
                
                cambios = False
                for idx in seleccion:
                    nombre_hoja, ruta_xml = self.hojas_disponibles[idx]
                    xml_path = os.path.join(tmpdir, ruta_xml)
                    
                    if os.path.exists(xml_path):
                        with open(xml_path, 'r', encoding='utf-8') as f:
                            content = f.read()

                        new_content = re.sub(r'<sheetProtection[^>]*>', '', content)
                        
                        if new_content != content:
                            with open(xml_path, 'w', encoding='utf-8') as f:
                                f.write(new_content)
                            hojas_ok.append(nombre_hoja)
                            cambios = True

                if not cambios:
                    messagebox.showinfo("Aviso", "No se detectó protección en las hojas seleccionadas.")
                    return

                ruta_final = nombre_base + "_desbloqueado.xlsx"
                with zipfile.ZipFile(ruta_final, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                    for root, dirs, files in os.walk(tmpdir):
                        for file in files:
                            fp = os.path.join(root, file)
                            rp = os.path.relpath(fp, tmpdir)
                            new_zip.write(fp, rp)
                
                res = f"¡Éxito!\n\nHojas liberadas:\n" + "\n".join(f"- {h}" for h in hojas_ok)
                res += f"\n\nGuardado: {os.path.basename(ruta_final)}"
                messagebox.showinfo("Completado", res)
                
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    w, h = 500, 520
    x = (root.winfo_screenwidth() // 2) - (w // 2)
    y = (root.winfo_screenheight() // 2) - (h // 2)
    root.geometry(f'{w}x{h}+{x}+{y}')
    app = ExcelUnlockerGUI(root)
    root.mainloop()
