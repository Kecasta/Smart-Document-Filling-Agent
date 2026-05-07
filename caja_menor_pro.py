import os
import sys
import subprocess
import json
import threading
import re

# --- PORTABLE BASE DIR ---
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, 'frozen', False) else os.path.expanduser("~"), "caja_menor_config.json")

def get_desktop_path():
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        desktop, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        return desktop
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

# --- AUTO-INSTALL DEPENDENCIES ---
def install_dependencies():
    deps = {'customtkinter': 'customtkinter', 'pandas': 'pandas', 'openpyxl': 'openpyxl',
            'win32com': 'pywin32', 'num2words': 'num2words', 'tkcalendar': 'tkcalendar'}
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

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

TARGET_PATH = os.path.join(get_desktop_path(), "recibos_de_caja.xlsx")


class CajaMenorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Logística Delfín S.A.S - Agente de Caja Menor")
        self.geometry("900x700")
        self.resizable(False, False)

        ico_path = os.path.join(BASE_DIR, "recibo.ico")
        if os.path.exists(ico_path):
            self.iconbitmap(ico_path)

        # Load persisted config
        cfg = load_config()
        self.template_path = ctk.StringVar(value=cfg.get("template_path", ""))
        self.source_paths = cfg.get("source_paths", [])

        self.tabview = ctk.CTkTabview(self, width=860, height=660)
        self.tabview.pack(padx=20, pady=20, fill="both", expand=True)

        self.tab_masivo = self.tabview.add("Procesamiento Masivo")
        self.tab_manual = self.tabview.add("Generador Manual")

        self.setup_tab_masivo()
        self.setup_tab_manual()

    # ─────────────────────────────── TAB MASIVO ───────────────────────────────
    def setup_tab_masivo(self):
        ctk.CTkLabel(self.tab_masivo, text="Automatización de Recibos Múltiples",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(10, 6))

        # --- Plantilla card ---
        card_t = ctk.CTkFrame(self.tab_masivo)
        card_t.pack(fill="x", padx=30, pady=6)
        ctk.CTkLabel(card_t, text="1 · Plantilla (Formato Excel)", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=12, pady=(8, 2))
        row_t = ctk.CTkFrame(card_t, fg_color="transparent")
        row_t.pack(fill="x", padx=12, pady=(0, 8))
        self.lbl_template = ctk.CTkLabel(row_t, text=self._short(self.template_path.get()), text_color="gray", anchor="w")
        self.lbl_template.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row_t, text="📂 Seleccionar Plantilla", width=180, command=self.cargar_plantilla).pack(side="right")

        # --- Fuente de datos card ---
        card_d = ctk.CTkFrame(self.tab_masivo)
        card_d.pack(fill="x", padx=30, pady=6)
        ctk.CTkLabel(card_d, text="2 · Archivos de Datos (Excel / CSV)", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=12, pady=(8, 2))
        row_d = ctk.CTkFrame(card_d, fg_color="transparent")
        row_d.pack(fill="x", padx=12, pady=(0, 6))
        self.lbl_sources = ctk.CTkLabel(row_d, text=self._sources_summary(), text_color="gray",
                                        anchor="w", wraplength=500, justify="left")
        self.lbl_sources.pack(side="left", fill="x", expand=True)
        btns = ctk.CTkFrame(row_d, fg_color="transparent")
        btns.pack(side="right")
        ctk.CTkButton(btns, text="Agregar Archivos", width=160, command=self.cargar_fuentes).pack(pady=2)
        ctk.CTkButton(btns, text="Limpiar Lista", width=160, fg_color="gray", hover_color="#555",
                      command=self.limpiar_fuentes).pack(pady=2)

        # --- Accion principal ---
        self.btn_generar_masivo = ctk.CTkButton(self.tab_masivo,
                                                text="Generar Recibos", height=48,
                                                fg_color="green", hover_color="darkgreen",
                                                font=ctk.CTkFont(size=15, weight="bold"),
                                                command=self.generar_masivo)
        self.btn_generar_masivo.pack(pady=18)

        self.lbl_status_masivo = ctk.CTkLabel(self.tab_masivo, text="", text_color="gray")
        self.lbl_status_masivo.pack(pady=4)

    # ─────────────────────────────── TAB MANUAL ───────────────────────────────
    def setup_tab_manual(self):
        ctk.CTkLabel(self.tab_manual, text="Generación de Recibo Único",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(10, 6))

        # Plantilla selector (shared)
        card_t = ctk.CTkFrame(self.tab_manual)
        card_t.pack(fill="x", padx=30, pady=6)
        row_t = ctk.CTkFrame(card_t, fg_color="transparent")
        row_t.pack(fill="x", padx=12, pady=8)
        self.lbl_template_manual = ctk.CTkLabel(row_t, textvariable=self.template_path, text_color="gray", anchor="w")
        self.lbl_template_manual.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row_t, text="📂 Plantilla", width=150, command=self.cargar_plantilla).pack(side="right")

        form_frame = ctk.CTkFrame(self.tab_manual, fg_color="transparent")
        form_frame.pack(pady=10, padx=50, fill="x")

        ctk.CTkLabel(form_frame, text="Fecha:").grid(row=0, column=0, padx=10, pady=8, sticky="w")
        self.entry_fecha = DateEntry(form_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entry_fecha.grid(row=0, column=1, padx=10, pady=8, sticky="w")

        ctk.CTkLabel(form_frame, text="Número de Recibo:").grid(row=1, column=0, padx=10, pady=8, sticky="w")
        self.entry_recibo = ctk.CTkEntry(form_frame, width=200)
        self.entry_recibo.grid(row=1, column=1, padx=10, pady=8, sticky="w")
        self.btn_refresh = ctk.CTkButton(form_frame, text="Actualizar N°", width=110, command=self.sugerir_numero_recibo)
        self.btn_refresh.grid(row=1, column=2, padx=10, pady=8)

        ctk.CTkLabel(form_frame, text="Valor numérico ($):").grid(row=2, column=0, padx=10, pady=8, sticky="w")
        self.entry_valor = ctk.CTkEntry(form_frame, width=200)
        self.entry_valor.grid(row=2, column=1, padx=10, pady=8, sticky="w")

        ctk.CTkLabel(form_frame, text="Pagado a (Beneficiario):").grid(row=3, column=0, padx=10, pady=8, sticky="w")
        self.entry_beneficiario = ctk.CTkEntry(form_frame, width=400)
        self.entry_beneficiario.grid(row=3, column=1, padx=10, pady=8, sticky="w")

        ctk.CTkLabel(form_frame, text="Concepto:").grid(row=4, column=0, padx=10, pady=8, sticky="w")
        self.entry_concepto = ctk.CTkEntry(form_frame, width=400)
        self.entry_concepto.grid(row=4, column=1, padx=10, pady=8, sticky="w")

        self.btn_generar_manual = ctk.CTkButton(self.tab_manual, text="Agregar e Imprimir Recibo Único",
                                                command=self.generar_manual, height=45,
                                                fg_color="darkblue", hover_color="#000080")
        self.btn_generar_manual.pack(pady=20)
        self.lbl_status_manual = ctk.CTkLabel(self.tab_manual, text="", text_color="gray")
        self.lbl_status_manual.pack(pady=4)

        self.sugerir_numero_recibo()

    # ─────────────────────────────── HELPERS ──────────────────────────────────
    def _short(self, path):
        if not path:
            return "⚠ No seleccionada"
        return f"✅ {os.path.basename(path)}"

    def _sources_summary(self):
        if not self.source_paths:
            return "Ningun archivo seleccionado"
        lines = [f"  {i+1}. {os.path.basename(p)}" for i, p in enumerate(self.source_paths)]
        header = f"{len(self.source_paths)} archivo(s) cargado(s):"
        return header + "\n" + "\n".join(lines)

    def _persist(self):
        save_config({"template_path": self.template_path.get(), "source_paths": self.source_paths})

    def _refresh_sources_ui(self):
        """Actualiza el label de archivos en la UI."""
        self.lbl_sources.configure(text=self._sources_summary())

    def cargar_plantilla(self):
        ruta = filedialog.askopenfilename(title="Seleccionar Plantilla Excel",
                                          filetypes=(("Excel", "*.xlsx *.xls"), ("Todos", "*.*")))
        if ruta:
            self.template_path.set(ruta)
            self.lbl_template.configure(text=self._short(ruta))
            self._persist()

    def cargar_fuentes(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar Archivos de Datos (multi-seleccion)",
            filetypes=(("Excel y CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*"))
        )
        if rutas:
            nuevas = [r for r in rutas if r not in self.source_paths]
            duplicadas = len(rutas) - len(nuevas)
            self.source_paths.extend(nuevas)
            self._refresh_sources_ui()
            self._persist()
            if duplicadas > 0:
                messagebox.showinfo("Info", f"Se ignoraron {duplicadas} archivo(s) ya existentes en la lista.")

    def limpiar_fuentes(self):
        self.source_paths = []
        self._refresh_sources_ui()
        self._persist()

    # ─────────────────────────────── MASTER INFO ──────────────────────────────
    def get_master_info(self):
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
                cell_val = ws.Cells(5 + i * 25, 4).Value
                cell_str = str(cell_val).strip() if cell_val is not None else ""
                if cell_val is None or cell_str == "" or "_________________________" in cell_str or cell_str == "Número de Recibo:":
                    next_block_idx = i
                    break
                if i > 500:
                    break
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

    # ─────────────────────────────── DATA PROCESSING ──────────────────────────
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
        if 'pago seguro' in desc_lower or 'pagos seguro' in desc_lower:
            concepto = "Pago Seguros Turbos"
        if 'finazauto' in desc_lower or 'finanzauto' in desc_lower:
            concepto = "Pagos Cuotas Turbo"
        return concepto, beneficiario

    def procesar_datos_masivos(self):
        if not self.source_paths:
            messagebox.showwarning("Advertencia", "Por favor selecciona archivos de datos primero.")
            return None
        dfs = []
        errores = []
        for ruta in self.source_paths:
            try:
                if ruta.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(ruta)
                elif ruta.endswith('.csv'):
                    df = pd.read_csv(ruta, encoding='utf-8', encoding_errors='replace')
                else:
                    continue
                df['__origen__'] = os.path.basename(ruta)  # Tag para trazabilidad
                dfs.append(df)
            except Exception as e:
                errores.append(f"{os.path.basename(ruta)}: {e}")
        if errores:
            messagebox.showerror("Error de lectura", "No se pudo leer:\n" + "\n".join(errores))
            if not dfs:
                return None
        if not dfs:
            return None

        # 1. Unir todos los DataFrames
        master_df = pd.concat(dfs, ignore_index=True)

        # 2. Eliminar filas exactamente iguales (duplicados reales entre archivos)
        cols_data = [c for c in master_df.columns if c != '__origen__']
        master_df = master_df.drop_duplicates(subset=cols_data).reset_index(drop=True)

        # 3. Filtros de exclusion
        str_cols = master_df.select_dtypes(include=['object', 'string']).columns
        mask = pd.Series(False, index=master_df.index)
        for col in str_cols:
            col_str = master_df[col].astype(str).str.lower()
            mask |= col_str.str.contains('transferencia', na=False)
            mask |= col_str.str.contains('seguridad social', na=False)
            mask |= col_str.str.contains('cuota de manejo', na=False)
            mask |= col_str.str.contains('cuotas de manejo', na=False)
        master_df = master_df[~mask]

        # 4. Sanitizacion y ordenamiento cronologico estricto (DD/MM/YYYY)
        fecha_col = next((c for c in master_df.columns if c.lower() == 'fecha'), None)
        if not fecha_col and 'Fecha' in master_df.columns:
            fecha_col = 'Fecha'
        if fecha_col:
            master_df[fecha_col] = pd.to_datetime(master_df[fecha_col], errors='coerce', dayfirst=True)
            # Eliminar filas con fechas nulas o corruptas (filtro de integridad)
            filas_antes = len(master_df)
            master_df = master_df.dropna(subset=[fecha_col]).reset_index(drop=True)
            filas_eliminadas = filas_antes - len(master_df)
            if filas_eliminadas > 0:
                print(f"Sanitizacion: {filas_eliminadas} fila(s) eliminadas por fecha invalida/nula.")
            # Orden ascendente estricto
            master_df = master_df.sort_values(by=fecha_col, ascending=True).reset_index(drop=True)
            master_df['Fecha'] = master_df[fecha_col].dt.strftime('%d/%m/%Y')

        # 5. Extraer concepto y beneficiario
        desc_col = next((c for c in master_df.columns if 'descrip' in c.lower()), None)
        desc = master_df[desc_col].fillna('').astype(str) if desc_col else pd.Series([''] * len(master_df))
        cb = desc.apply(self.parse_descripcion)
        master_df['Concepto'] = [x[0] for x in cb]
        master_df['Beneficiario'] = [x[1] for x in cb]

        return master_df

    # ─────────────────────────────── FILL EXCEL ───────────────────────────────
    def llenar_datos_com(self, ws, start_row, datos):
        fecha = datos.get('Fecha', '')
        recibo = datos.get('Numero_Recibo', '')
        recibo_str = str(recibo).zfill(4) if str(recibo).isdigit() else str(recibo)
        beneficiario = datos.get('Beneficiario', '')
        try:
            valor = float(datos.get('Valor', 0))
            if pd.isna(valor):
                valor = 0
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

    # ─────────────────────────────── MASIVO ───────────────────────────────────
    def generar_masivo(self):
        if not WIN32_AVAILABLE:
            messagebox.showerror("Error", "La librería win32com no está disponible.")
            return
        tpl = self.template_path.get()
        if not tpl or not os.path.exists(tpl):
            messagebox.showerror("Error", "Selecciona una plantilla Excel válida primero (Paso 1).")
            return
        df = self.procesar_datos_masivos()
        if df is None or df.empty:
            messagebox.showinfo("Información", "No se encontraron datos válidos para procesar.")
            return
        self.lbl_status_masivo.configure(text="Procesando... (Por favor espere)", text_color="gray")
        self.btn_generar_masivo.configure(state="disabled")
        self.update()
        threading.Thread(target=self._thread_generar_masivo, args=(df, tpl), daemon=True).start()

    def _thread_generar_masivo(self, df, tpl):
        pythoncom.CoInitialize()
        try:
            if not os.path.exists(TARGET_PATH):
                import shutil
                shutil.copy(tpl, TARGET_PATH)
            next_block_idx, max_recibo = self.get_master_info()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(TARGET_PATH)
            ws = wb.ActiveSheet
            template_rows = 25
            for index, row in df.iterrows():
                row['Numero_Recibo'] = max_recibo + 1 + index
                start_row = 1 + (next_block_idx + index) * template_rows
                if start_row > 1:
                    ws.Rows(f"1:{template_rows}").Copy(ws.Rows(f"{start_row}:{start_row + template_rows - 1}"))
                self.llenar_datos_com(ws, start_row, row)
            wb.Save()
            wb.Close(False)
            self.lbl_status_masivo.configure(text=f"✅ {len(df)} recibos generados exitosamente", text_color="green")
            messagebox.showinfo("Éxito", f"Se generaron {len(df)} recibos:\n{TARGET_PATH}")
            os.startfile(TARGET_PATH)
        except Exception as e:
            self.lbl_status_masivo.configure(text="❌ Error en generación", text_color="red")
            messagebox.showerror("Error", f"Ocurrió un error: {e}")
        finally:
            try:
                excel.Quit()
            except:
                pass
            pythoncom.CoUninitialize()
            self.btn_generar_masivo.configure(state="normal")
        self.sugerir_numero_recibo()

    # ─────────────────────────────── MANUAL ───────────────────────────────────
    def generar_manual(self):
        if not WIN32_AVAILABLE:
            messagebox.showerror("Error", "La librería win32com no está disponible.")
            return
        tpl = self.template_path.get()
        if not tpl or not os.path.exists(tpl):
            messagebox.showerror("Error", "Selecciona una plantilla Excel válida primero.")
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
        self.lbl_status_manual.configure(text="Generando...", text_color="gray")
        self.update()
        datos = {'Fecha': fecha, 'Numero_Recibo': recibo, 'Beneficiario': beneficiario,
                 'Valor': valor, 'Concepto': concepto}
        pdf_path = os.path.join(get_desktop_path(), f"Recibo_Caja_Menor_{recibo}.pdf")
        try:
            if not os.path.exists(TARGET_PATH):
                import shutil
                shutil.copy(tpl, TARGET_PATH)
            next_block_idx, _ = self.get_master_info()
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
            wb.Save()
            ws.PageSetup.PrintArea = f"$A${start_row}:$F${start_row + 24}"
            ws.ExportAsFixedFormat(0, pdf_path)
            ws.PageSetup.PrintArea = ""
            wb.Save()
            wb.Close(False)
            self.lbl_status_manual.configure(text="✅ Recibo generado y PDF listo", text_color="green")
            os.startfile(pdf_path)
        except Exception as e:
            self.lbl_status_manual.configure(text="❌ Error en generación", text_color="red")
            messagebox.showerror("Error", f"Error: {e}")
        finally:
            try:
                excel.Quit()
            except:
                pass
        self.sugerir_numero_recibo()


if __name__ == "__main__":
    app = CajaMenorApp()
    app.mainloop()
