import os
import sys
import subprocess
import datetime
import threading
import re

# --- PORTABLE BASE DIR ---
# Cuando está empaquetado como .exe, sys.frozen = True y los recursos
# están en sys._MEIPASS. En desarrollo, usa el directorio del script.
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_desktop_path():
    """Lee la ruta real del Escritorio desde el registro de Windows.
    Funciona aunque el Escritorio esté redirigido a OneDrive."""
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
        )
        desktop, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        return desktop
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")

# --- AUTO-INSTALL DEPENDENCIES ---
def install_dependencies():
    deps = {
        'customtkinter': 'customtkinter',
        'pandas': 'pandas',
        'openpyxl': 'openpyxl',
        'win32com': 'pywin32',
        'num2words': 'num2words',
        'tkcalendar': 'tkcalendar'
    }
    
    for module, package in deps.items():
        try:
            if module != 'win32com':
                __import__(module)
        except ImportError:
            print(f"Instalando {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_dependencies()

import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from num2words import num2words
from tkcalendar import DateEntry

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# --- CONFIGURATION ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Plantilla de muestra (viaja junto al script o al .exe)
TEMPLATE_PATH = os.path.join(BASE_DIR, "Formato_Caja_Menor_Logistica_Delfin.xlsx")

# Archivo maestro: se detecta el Escritorio real (OneDrive o local)
TARGET_PATH = os.path.join(get_desktop_path(), "recibos_de_caja.xlsx")

class CajaMenorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Logística Delfín S.A.S - Agente de Caja Menor")
        self.geometry("850x650")
        self.resizable(False, False)
        
        ico_path = os.path.join(BASE_DIR, "recibo.ico")
        if os.path.exists(ico_path):
            self.iconbitmap(ico_path)
        
        import shutil
        if not os.path.exists(TARGET_PATH):
            if os.path.exists(TEMPLATE_PATH):
                shutil.copy(TEMPLATE_PATH, TARGET_PATH)
            else:
                messagebox.showerror("Error", f"No se encontró la plantilla de muestra: {TEMPLATE_PATH}")
        elif not os.path.exists(TEMPLATE_PATH):
            messagebox.showerror("Error", f"Falta la plantilla de muestra: {TEMPLATE_PATH}")
        
        self.tabview = ctk.CTkTabview(self, width=800, height=600)
        self.tabview.pack(padx=20, pady=20, fill="both", expand=True)
        
        self.tab_masivo = self.tabview.add("Procesamiento Masivo")
        self.tab_manual = self.tabview.add("Generador Manual")
        
        self.setup_tab_masivo()
        self.setup_tab_manual()
        
        self.archivos_seleccionados = []
        
    def setup_tab_masivo(self):
        self.lbl_masivo = ctk.CTkLabel(self.tab_masivo, text="Automatización de Recibos Múltiples", font=ctk.CTkFont(size=20, weight="bold"))
        self.lbl_masivo.pack(pady=(10, 10))
        
        self.btn_auto = ctk.CTkButton(self.tab_masivo, text="Proceso Automático (Cuenta Bancolombia + Caja Menor)", 
                                      command=self.procesar_automatico, height=45, fg_color="#b8860b", hover_color="#8b6508", font=ctk.CTkFont(weight="bold"))
        self.btn_auto.pack(pady=10)
        
        ctk.CTkLabel(self.tab_masivo, text="--- O ---", text_color="gray").pack(pady=5)
        
        self.btn_seleccionar = ctk.CTkButton(self.tab_masivo, text="Agregar Archivos Personalizados (.xlsx, .csv)", command=self.seleccionar_archivos, height=35)
        self.btn_seleccionar.pack(pady=10)
        
        self.txt_archivos = ctk.CTkTextbox(self.tab_masivo, width=650, height=80)
        self.txt_archivos.pack(pady=10)
        self.txt_archivos.insert("0.0", "Ningún archivo seleccionado.")
        self.txt_archivos.configure(state="disabled")
        
        self.btn_generar_masivo = ctk.CTkButton(self.tab_masivo, text="Agregar al Formato Original", command=self.generar_masivo, height=45, fg_color="green", hover_color="darkgreen")
        self.btn_generar_masivo.pack(pady=20)
        
        self.lbl_status_masivo = ctk.CTkLabel(self.tab_masivo, text="", text_color="gray")
        self.lbl_status_masivo.pack(pady=5)

    def setup_tab_manual(self):
        self.lbl_manual = ctk.CTkLabel(self.tab_manual, text="Generación de Recibo Único", font=ctk.CTkFont(size=20, weight="bold"))
        self.lbl_manual.pack(pady=(10, 20))
        
        form_frame = ctk.CTkFrame(self.tab_manual, fg_color="transparent")
        form_frame.pack(pady=10, padx=50, fill="x")
        
        ctk.CTkLabel(form_frame, text="Fecha:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.entry_fecha = DateEntry(form_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entry_fecha.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        ctk.CTkLabel(form_frame, text="Número de Recibo:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.entry_recibo = ctk.CTkEntry(form_frame, width=200)
        self.entry_recibo.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        
        # Sugerir recibo basado en el archivo maestro
        self.btn_refresh = ctk.CTkButton(form_frame, text="Actualizar N°", width=100, command=self.sugerir_numero_recibo)
        self.btn_refresh.grid(row=1, column=2, padx=10, pady=10)
        
        ctk.CTkLabel(form_frame, text="Valor numérico ($):").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.entry_valor = ctk.CTkEntry(form_frame, width=200)
        self.entry_valor.grid(row=2, column=1, padx=10, pady=10, sticky="w")
        
        ctk.CTkLabel(form_frame, text="Pagado a (Beneficiario):").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.entry_beneficiario = ctk.CTkEntry(form_frame, width=400)
        self.entry_beneficiario.grid(row=3, column=1, padx=10, pady=10, sticky="w")
        
        ctk.CTkLabel(form_frame, text="Concepto:").grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.entry_concepto = ctk.CTkEntry(form_frame, width=400)
        self.entry_concepto.grid(row=4, column=1, padx=10, pady=10, sticky="w")
        
        self.btn_generar_manual = ctk.CTkButton(self.tab_manual, text="Agregar e Imprimir Recibo Único", command=self.generar_manual, height=45, fg_color="darkblue", hover_color="#000080")
        self.btn_generar_manual.pack(pady=30)
        
        self.lbl_status_manual = ctk.CTkLabel(self.tab_manual, text="", text_color="gray")
        self.lbl_status_manual.pack(pady=5)
        
        # Llamar sugerencia inicialmente de forma segura (sincrónicamente)
        self.sugerir_numero_recibo()

    def get_master_info(self):
        """Escanea el archivo Formato_Caja_Menor_Logistica_Delfin.xlsx para saber dónde seguir escribiendo"""
        if not WIN32_AVAILABLE:
            return 0, 0
            
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(TARGET_PATH)
            ws = wb.ActiveSheet
            
            max_recibo = 0
            next_block_idx = 0
            i = 0
            
            while True:
                # Cada bloque mide 25 filas. La celda del N de recibo es la fila 5 relativa.
                cell_val = ws.Cells(5 + i * 25, 4).Value
                cell_str = str(cell_val).strip() if cell_val is not None else ""
                
                if cell_val is None or cell_str == "" or "_________________________" in cell_str or cell_str == "Número de Recibo:":
                    next_block_idx = i
                    break
                    
                if i > 500: # Salvaguarda contra bucles infinitos
                    break
                    
                # Extraer número
                nums = re.findall(r'\d+', cell_str)
                if nums:
                    num = int(nums[-1])
                    if num > max_recibo:
                        max_recibo = num
                
                i += 1
                
            wb.Close(False)
            return next_block_idx, max_recibo
        except Exception as e:
            print("Error leyendo maestro:", e)
            try:
                wb.Close(False)
            except:
                pass
            return 0, 0
        finally:
            try:
                excel.Quit()
            except:
                pass
            pythoncom.CoUninitialize()

    def sugerir_numero_recibo(self):
        self.btn_refresh.configure(state="disabled")
        try:
            _, max_recibo = self.get_master_info()
            self.entry_recibo.delete(0, 'end')
            self.entry_recibo.insert(0, str(max_recibo + 1).zfill(4))
        finally:
            self.btn_refresh.configure(state="normal")

    def seleccionar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar Archivos de Datos",
            filetypes=(("Excel y CSV", "*.xlsx *.csv"), ("Todos los archivos", "*.*"))
        )
        if rutas:
            self.archivos_seleccionados = list(rutas)
            self.txt_archivos.configure(state="normal")
            self.txt_archivos.delete("0.0", "end")
            self.txt_archivos.insert("0.0", "\n".join(self.archivos_seleccionados))
            self.txt_archivos.configure(state="disabled")

    def procesar_automatico(self):
        archivos = os.listdir('.')
        archivos_bancolombia = [os.path.abspath(f) for f in archivos if 'cuenta_bancolombia' in f.lower() and f.endswith('.xlsx') and not f.startswith('~') and 'formato' not in f.lower() and 'reporte' not in f.lower() and 'recibo' not in f.lower()]
        archivos_caja = [os.path.abspath(f) for f in archivos if 'caja_menor' in f.lower() and f.endswith('.xlsx') and not f.startswith('~') and 'formato' not in f.lower() and 'reporte' not in f.lower() and 'recibo' not in f.lower()]
        
        rutas = archivos_bancolombia + archivos_caja
        if not rutas:
            messagebox.showerror("Error", "No se encontraron archivos de 'cuenta_bancolombia' o 'caja_menor' en la carpeta raíz.")
            return
            
        self.archivos_seleccionados = rutas
        self.txt_archivos.configure(state="normal")
        self.txt_archivos.delete("0.0", "end")
        self.txt_archivos.insert("0.0", "\n".join(self.archivos_seleccionados))
        self.txt_archivos.configure(state="disabled")
        
        self.generar_masivo()

    def parse_descripcion(self, desc):
        desc_clean = str(desc).strip()
        desc_lower = desc_clean.lower()
        words = desc_clean.split()
        
        if len(words) >= 3:
            concepto = " ".join(words[:2]).capitalize()
            beneficiario = " ".join(words[2:]).title()
        elif len(words) == 2:
            concepto = words[0].capitalize()
            beneficiario = words[1].title()
        else:
            concepto = desc_clean.capitalize()
            beneficiario = "Portador"
            
        # Reglas de negocio para sobreescribir concepto
        if 'pago seguro' in desc_lower or 'pagos seguro' in desc_lower:
            concepto = "Pago Seguros Turbos"
        if 'finazauto' in desc_lower or 'finanzauto' in desc_lower:
            concepto = "Pagos Cuotas Turbo"
            
        return concepto, beneficiario

    def procesar_datos_masivos(self):
        if not self.archivos_seleccionados:
            messagebox.showwarning("Advertencia", "Por favor selecciona archivos primero.")
            return None
        
        dfs = []
        for ruta in self.archivos_seleccionados:
            try:
                if ruta.endswith('.xlsx') or ruta.endswith('.xls'):
                    df = pd.read_excel(ruta)
                elif ruta.endswith('.csv'):
                    df = pd.read_csv(ruta)
                dfs.append(df)
            except Exception as e:
                messagebox.showerror("Error", f"Error al leer {os.path.basename(ruta)}: {str(e)}")
                return None
                
        if not dfs:
            return None
            
        master_df = pd.concat(dfs, ignore_index=True)
        
        str_cols = master_df.select_dtypes(include=['object', 'string']).columns
        mask = pd.Series(False, index=master_df.index)
        for col in str_cols:
            col_str = master_df[col].astype(str).str.lower()
            mask = mask | col_str.str.contains('transferencia', na=False)
            mask = mask | col_str.str.contains('seguridad social', na=False)
            mask = mask | col_str.str.contains('cuota de manejo', na=False)
            mask = mask | col_str.str.contains('cuotas de manejo', na=False)
        master_df = master_df[~mask]
        
        if 'Fecha' in master_df.columns:
            master_df['Fecha'] = pd.to_datetime(master_df['Fecha'], errors='coerce')
            master_df = master_df.sort_values(by='Fecha').reset_index(drop=True)
            master_df['Fecha'] = master_df['Fecha'].dt.strftime('%d/%m/%Y').fillna('')
            
        desc_col = next((c for c in master_df.columns if 'descrip' in c.lower()), None)
        if desc_col:
            desc = master_df[desc_col].fillna('').astype(str)
        else:
            desc = pd.Series(['']*len(master_df))
        
        conceptos_beneficiarios = desc.apply(self.parse_descripcion)
        master_df['Concepto'] = [cb[0] for cb in conceptos_beneficiarios]
        master_df['Beneficiario'] = [cb[1] for cb in conceptos_beneficiarios]
            
        return master_df

    def llenar_datos_com(self, ws, start_row, datos):
        fecha = datos.get('Fecha', '')
        recibo = datos.get('Numero_Recibo', '')
        recibo_str = str(recibo).zfill(4) if str(recibo).isdigit() else str(recibo)
        beneficiario = datos.get('Beneficiario', '')
        
        try:
            valor = float(datos.get('Valor', 0))
            if pd.isna(valor): valor = 0
        except Exception:
            valor = 0
            
        concepto = datos.get('Concepto', '')
        
        try:
            valor_letras = num2words(int(valor), lang='es').upper()
        except Exception:
            valor_letras = str(valor)
            
        ws.Cells(start_row + 4, 1).Value = f"Fecha: {fecha}"
        ws.Cells(start_row + 4, 4).Value = f"Número de Recibo: {recibo_str}"
        ws.Cells(start_row + 6, 1).Value = f"pagado a: {beneficiario}"
        ws.Cells(start_row + 6, 4).Value = "valor: $"
        ws.Cells(start_row + 6, 5).Value = f"{valor:,.2f}"
        ws.Cells(start_row + 9, 1).Value = f"Concepto: {concepto}"
        ws.Cells(start_row + 13, 1).Value = f"Valor (en letras): {valor_letras} PESOS M/CTE."

    def generar_masivo(self):
        if not WIN32_AVAILABLE:
            messagebox.showerror("Error", "La librería win32com no está disponible.")
            return

        df = self.procesar_datos_masivos()
        if df is None or df.empty:
            messagebox.showinfo("Información", "No se encontraron datos válidos para procesar.")
            return
            
        self.lbl_status_masivo.configure(text="Analizando formato maestro y agregando recibos... (Por favor espere)")
        self.btn_generar_masivo.configure(state="disabled")
        self.update()
        
        # Ejecutar en segundo plano para no congelar la UI
        threading.Thread(target=self._thread_generar_masivo, args=(df,), daemon=True).start()

    def _thread_generar_masivo(self, df):
        pythoncom.CoInitialize()
        try:
            next_block_idx, max_recibo = self.get_master_info()
            
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(TARGET_PATH)
            ws = wb.ActiveSheet
            
            template_rows = 25
            
            for index, row in df.iterrows():
                # Asignar número consecutivo al vuelo
                row['Numero_Recibo'] = max_recibo + 1 + index
                
                start_row = 1 + (next_block_idx + index) * template_rows
                if start_row > 1:
                    # Copiar siempre el primer bloque (filas 1 a 25)
                    ws.Rows(f"1:{template_rows}").Copy(ws.Rows(f"{start_row}:{start_row + template_rows - 1}"))
                    
                self.llenar_datos_com(ws, start_row, row)
                
            # Sobrescribir el archivo original
            wb.Save()
            wb.Close(False)
            
            self.lbl_status_masivo.configure(text=f"Agregados exitosamente", text_color="green")
            messagebox.showinfo("Éxito", f"Se agregaron {len(df)} recibos al archivo del escritorio:\n{TARGET_PATH}")
            os.startfile(TARGET_PATH)
            
        except Exception as e:
            self.lbl_status_masivo.configure(text="Error en generación", text_color="red")
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")
        finally:
            try:
                excel.Quit()
            except:
                pass
            pythoncom.CoUninitialize()
            self.btn_generar_masivo.configure(state="normal")
        
        # Actualizar la sugerencia de la tab manual
        self.sugerir_numero_recibo()

    def generar_manual(self):
        if not WIN32_AVAILABLE:
            messagebox.showerror("Error", "La librería win32com no está disponible.")
            return

        fecha = self.entry_fecha.get()
        recibo = self.entry_recibo.get()
        valor_str = self.entry_valor.get().replace(',', '').replace('$', '').strip()
        beneficiario = self.entry_beneficiario.get()
        concepto = self.entry_concepto.get()
        
        if not all([fecha, recibo, valor_str, beneficiario, concepto]):
            messagebox.showwarning("Advertencia", "Por favor complete todos los campos.")
            return
            
        try:
            valor = float(valor_str)
        except ValueError:
            messagebox.showerror("Error", "El valor debe ser numérico.")
            return
            
        self.lbl_status_manual.configure(text="Agregando al maestro y generando PDF...")
        self.update()
        
        datos = {
            'Fecha': fecha,
            'Numero_Recibo': recibo,
            'Beneficiario': beneficiario,
            'Valor': valor,
            'Concepto': concepto
        }
        
        pdf_path = os.path.abspath(f"Recibo_Caja_Menor_{recibo}.pdf")
        
        try:
            next_block_idx, max_recibo = self.get_master_info()
            
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(TARGET_PATH)
            ws = wb.ActiveSheet
            
            template_rows = 25
            start_row = 1 + next_block_idx * template_rows
            
            if start_row > 1:
                ws.Rows(f"1:{template_rows}").Copy(ws.Rows(f"{start_row}:{start_row + template_rows - 1}"))
                
            self.llenar_datos_com(ws, start_row, datos)
            
            # Guardar maestro
            wb.Save()
            
            # Exportar SÓLO este recibo a PDF
            ws.PageSetup.PrintArea = f"$A${start_row}:$F${start_row + 24}"
            ws.ExportAsFixedFormat(0, pdf_path)
            ws.PageSetup.PrintArea = ""
            wb.Save() # Para quitar el area de impresion guardada
            wb.Close(False)
                
            self.lbl_status_manual.configure(text=f"Agregado e impreso", text_color="green")
            os.startfile(pdf_path)
            
        except Exception as e:
            self.lbl_status_manual.configure(text="Error en generación", text_color="red")
            messagebox.showerror("Error", f"Error: {str(e)}")
        finally:
            try:
                excel.Quit()
            except:
                pass
        
        # Actualizar
        self.sugerir_numero_recibo()

if __name__ == "__main__":
    app = CajaMenorApp()
    app.mainloop()
