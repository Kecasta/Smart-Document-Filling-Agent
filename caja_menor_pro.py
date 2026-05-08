"""
Smart-Document-Filling-Agent v2.0
Desarrollado por: Ing. Kevin Seryeit Castañeda Aldana
Año: 2026
"""
__version__ = "2.0.1"
__author__ = "Ing. Kevin Seryeit Castañeda Aldana"

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
            'win32com': 'pywin32', 'num2words': 'num2words', 'tkcalendar': 'tkcalendar',
            'requests': 'requests'}
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
import requests
from datetime import datetime, timedelta
import winreg
import sqlite3
import hashlib

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

TARGET_PATH = os.path.join(get_desktop_path(), "recibos_de_caja.xlsx")

# --- TRIAL CONFIG ---
TRIAL_DAYS = 15
REG_PATH = r"Software\Classes\CLSID\{B54F3741-5B07-4A96-971D-0128C0A31048}"

def get_network_time():
    """Obtiene la fecha actual de internet para evitar manipulacion del reloj local."""
    try:
        # Intentar con WorldTimeAPI
        response = requests.get("http://worldtimeapi.org/api/timezone/Etc/UTC", timeout=5)
        if response.status_code == 200:
            data = response.json()
            return datetime.fromisoformat(data['datetime'].replace('Z', '+00:00')).replace(tzinfo=None)
    except:
        pass
    try:
        # Fallback: Usar el header Date de Google
        response = requests.head("http://www.google.com", timeout=5)
        date_str = response.headers.get('Date')
        if date_str:
            # Formato: "Fri, 08 May 2026 16:15:00 GMT"
            import email.utils
            return datetime(*email.utils.parsedate(date_str)[:6])
    except:
        pass
    return datetime.now()

def check_trial_status():
    """Verifica el estado del periodo de prueba en el Registro de Windows."""
    now = get_network_time()
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH)
        try:
            val, _ = winreg.QueryValueEx(key, "Installed")
            start_date = datetime.fromtimestamp(float(val))
        except FileNotFoundError:
            # Primera ejecucion
            start_date = now
            winreg.SetValueEx(key, "Installed", 0, winreg.REG_SZ, str(start_date.timestamp()))
            winreg.SetValueEx(key, "LastRun", 0, winreg.REG_SZ, str(now.timestamp()))
        
        # Validar si el usuario atraso el reloj
        try:
            last_val, _ = winreg.QueryValueEx(key, "LastRun")
            last_run = datetime.fromtimestamp(float(last_val))
            if now < last_run:
                # El tiempo actual es anterior al ultimo uso registrado (sospechoso)
                now = last_run
            else:
                winreg.SetValueEx(key, "LastRun", 0, winreg.REG_SZ, str(now.timestamp()))
        except:
            pass
            
        winreg.CloseKey(key)
        
        days_passed = (now - start_date).days
        days_left = max(0, TRIAL_DAYS - days_passed)
        is_expired = days_passed >= TRIAL_DAYS
        
        return is_expired, days_left
    except Exception as e:
        print(f"Trial Check Error: {e}")
        return False, TRIAL_DAYS

HISTORY_DB = os.path.join(os.path.dirname(CONFIG_FILE), "data_history.db")

class HistoryManager:
    def __init__(self):
        self.conn = sqlite3.connect(HISTORY_DB, check_same_thread=False)
        self._create_table()

    def _create_table(self):
        with self.conn:
            self.conn.execute("CREATE TABLE IF NOT EXISTS processed_records (record_hash TEXT PRIMARY KEY)")

    def is_processed(self, record_hash):
        cur = self.conn.cursor()
        cur.execute("SELECT 1 FROM processed_records WHERE record_hash = ?", (record_hash,))
        return cur.fetchone() is not None

    def add_records(self, hashes):
        with self.conn:
            self.conn.executemany("INSERT OR IGNORE INTO processed_records (record_hash) VALUES (?)", [(h,) for h in hashes])

    def reset_history(self):
        with self.conn:
            self.conn.execute("DELETE FROM processed_records")

    @staticmethod
    def generate_hash(row):
        """Genera una huella digital para el registro basada en sus campos clave."""
        raw_str = f"{row.get('Fecha','')}|{row.get('Beneficiario','')}|{row.get('Valor',0)}|{row.get('Concepto','')}"
        return hashlib.sha256(raw_str.encode('utf-8', errors='replace')).hexdigest()


class CajaMenorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Smart-Document-Filling-Agent v2.0 | Ing. Kevin Seryeit Castañeda Aldana")
        self.geometry("900x750")  # Aumentado un poco para el banner
        self.resizable(False, False)

        ico_path = os.path.join(BASE_DIR, "recibo.ico")
        if os.path.exists(ico_path):
            self.iconbitmap(ico_path)

        # Trial Verification
        self.is_expired, self.days_left = check_trial_status()

        # History Manager
        self.history = HistoryManager()

        # Load persisted config
        cfg = load_config()
        self.template_path = ctk.StringVar(value=cfg.get("template_path", ""))
        self.source_paths = cfg.get("source_paths", [])

        # --- Trial Banner ---
        if self.is_expired:
            self.banner = ctk.CTkFrame(self, fg_color="#8B0000", height=40)
            self.banner.pack(fill="x", side="top")
            ctk.CTkLabel(self.banner, text="🚨 PERIODO DE EVALUACIÓN FINALIZADO", 
                         font=ctk.CTkFont(weight="bold"), text_color="white").pack(pady=5)
            self.contact_lbl = ctk.CTkLabel(self, text="Contactar a Kevin Castañeda Aldana para activar licencia permanente",
                                            text_color="red", font=ctk.CTkFont(size=12, slant="italic"))
            self.contact_lbl.pack(side="bottom", pady=10)
        else:
            self.lbl_trial = ctk.CTkLabel(self, text=f"Días restantes de prueba: {self.days_left}",
                                          font=ctk.CTkFont(size=11), text_color="#555")
            self.lbl_trial.place(relx=0.98, rely=0.01, anchor="ne")

        self.tabview = ctk.CTkTabview(self, width=860, height=620)
        self.tabview.pack(padx=20, pady=(10, 20), fill="both", expand=True)

        self.tab_masivo = self.tabview.add("Procesamiento Masivo")
        self.tab_manual = self.tabview.add("Generador Manual")

        self.setup_tab_masivo()
        self.setup_tab_manual()

        if self.is_expired:
            self._disable_all_actions()

        # --- Branding Footer ---
        self.footer = ctk.CTkLabel(self, text="Desarrollado por Ing. Kevin Seryeit Castañeda Aldana - 2026",
                                   font=ctk.CTkFont(size=10), text_color="gray")
        self.footer.pack(side="bottom", pady=5)

    def _disable_all_actions(self):
        """Desactiva los botones principales si el periodo expiro."""
        messagebox.showwarning("Trial Expirado", "El periodo de evaluación de 15 días ha finalizado.\n\nPor favor contacte al desarrollador para obtener la versión completa.")
        # Se desactivan los botones via setup o aqui directamente si ya existen
        pass

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
        self.btn_sel_template = ctk.CTkButton(row_t, text="📂 Seleccionar Plantilla", width=180, 
                                               command=self.cargar_plantilla,
                                               state="normal" if not self.is_expired else "disabled")
        self.btn_sel_template.pack(side="right")

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
        self.btn_add_files = ctk.CTkButton(btns, text="Agregar Archivos", width=160, 
                                            command=self.cargar_fuentes,
                                            state="normal" if not self.is_expired else "disabled")
        self.btn_add_files.pack(pady=2)
        self.btn_clear_files = ctk.CTkButton(btns, text="Limpiar Lista", width=160, fg_color="gray", hover_color="#555",
                                              command=self.limpiar_fuentes,
                                              state="normal" if not self.is_expired else "disabled")
        self.btn_clear_files.pack(pady=2)

        # --- Accion principal ---
        self.btn_generar_masivo = ctk.CTkButton(self.tab_masivo,
                                                text="Generar Recibos", height=48,
                                                fg_color="green" if not self.is_expired else "gray", 
                                                hover_color="darkgreen" if not self.is_expired else "gray",
                                                font=ctk.CTkFont(size=15, weight="bold"),
                                                state="normal" if not self.is_expired else "disabled",
                                                command=self.generar_masivo)
        self.btn_generar_masivo.pack(pady=14)

        self.lbl_status_masivo = ctk.CTkLabel(self.tab_masivo, text="", text_color="gray")
        self.lbl_status_masivo.pack(pady=2)

        # --- Boton Reset Historial ---
        self.btn_reset_hist = ctk.CTkButton(self.tab_masivo, text="Reiniciar Historial de Duplicados", 
                                            fg_color="transparent", text_color="gray", 
                                            hover_color="#333", font=ctk.CTkFont(size=10),
                                            width=180, height=20,
                                            command=self.confirmar_reset_historial)
        self.btn_reset_hist.pack(side="bottom", pady=5)

    def confirmar_reset_historial(self):
        if messagebox.askyesno("Confirmar", "¿Deseas borrar permanentemente el historial de registros procesados?\n\nEsto permitira volver a procesar registros antiguos."):
            self.history.reset_history()
            messagebox.showinfo("Hecho", "Historial de duplicados vaciado.")

        # --- Debug Log ---
        ctk.CTkLabel(self.tab_masivo, text="Log de Procesamiento:",
                     font=ctk.CTkFont(size=11, weight="bold"), anchor="w").pack(anchor="w", padx=32, pady=(6, 0))
        self.txt_log = ctk.CTkTextbox(self.tab_masivo, width=800, height=100, font=ctk.CTkFont(size=11))
        self.txt_log.pack(padx=30, pady=(2, 8))
        self.txt_log.configure(state="disabled")

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
        self.btn_tpl_manual = ctk.CTkButton(row_t, text="📂 Plantilla", width=150, 
                                             command=self.cargar_plantilla,
                                             state="normal" if not self.is_expired else "disabled")
        self.btn_tpl_manual.pack(side="right")

        form_frame = ctk.CTkFrame(self.tab_manual, fg_color="transparent")
        form_frame.pack(pady=10, padx=50, fill="x")

        ctk.CTkLabel(form_frame, text="Fecha:").grid(row=0, column=0, padx=10, pady=8, sticky="w")
        self.entry_fecha = DateEntry(form_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entry_fecha.grid(row=0, column=1, padx=10, pady=8, sticky="w")

        ctk.CTkLabel(form_frame, text="Número de Recibo:").grid(row=1, column=0, padx=10, pady=8, sticky="w")
        self.entry_recibo = ctk.CTkEntry(form_frame, width=200)
        self.entry_recibo.grid(row=1, column=1, padx=10, pady=8, sticky="w")
        self.btn_refresh = ctk.CTkButton(form_frame, text="Actualizar N°", width=110, 
                                          command=self.sugerir_numero_recibo,
                                          state="normal" if not self.is_expired else "disabled")
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
                                                fg_color="darkblue" if not self.is_expired else "gray", 
                                                hover_color="#000080" if not self.is_expired else "gray",
                                                state="normal" if not self.is_expired else "disabled")
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
    def _log(self, msg):
        """Escribe una linea en el panel de debug."""
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def parse_descripcion(self, desc):
        """Extrae concepto y beneficiario de forma generica (sin reglas de negocio)."""
        desc_clean = str(desc).strip()
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
        return concepto, beneficiario

    def procesar_datos_masivos(self):
        if not self.source_paths:
            messagebox.showwarning("Advertencia", "Por favor selecciona archivos de datos primero.")
            return None

        # Limpiar log
        self.txt_log.configure(state="normal")
        self.txt_log.delete("0.0", "end")
        self.txt_log.configure(state="disabled")

        archivos_nombres = [os.path.basename(p) for p in self.source_paths]
        self._log(f"Archivos detectados: {archivos_nombres}")

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
                self._log(f"  -> {os.path.basename(ruta)}: {len(df)} filas leidas")
                df['__origen__'] = os.path.basename(ruta)
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
        total_bruto = len(master_df)

        # 2. Eliminar duplicados exactos entre archivos
        cols_data = [c for c in master_df.columns if c != '__origen__']
        master_df = master_df.drop_duplicates(subset=cols_data).reset_index(drop=True)
        dupes = total_bruto - len(master_df)
        if dupes > 0:
            self._log(f"Duplicados eliminados: {dupes}")

        # 3. Normalizar valores numericos (columna Valor)
        valor_col = next((c for c in master_df.columns if c.lower() == 'valor'), None)
        if valor_col:
            master_df[valor_col] = pd.to_numeric(
                master_df[valor_col].astype(str).str.replace(r'[^\d.\-]', '', regex=True),
                errors='coerce'
            ).fillna(0)

        # 4. Sanitizacion de fechas y ordenamiento cronologico
        #    Formato fuente: YYYY-MM-DD (ISO 8601) — pandas lo detecta nativamente
        fecha_col = next((c for c in master_df.columns if c.lower() == 'fecha'), None)
        if fecha_col:
            raw_sample = str(master_df[fecha_col].dropna().iloc[0]) if master_df[fecha_col].notna().any() else ''
            self._log(f"Formato de fecha detectado (muestra): {raw_sample}")
            master_df[fecha_col] = pd.to_datetime(master_df[fecha_col], errors='coerce')
            nulas = master_df[fecha_col].isna().sum()
            if nulas > 0:
                self._log(f"Fechas invalidas/nulas descartadas: {nulas}")
                master_df = master_df.dropna(subset=[fecha_col]).reset_index(drop=True)
            master_df = master_df.sort_values(by=fecha_col, ascending=True).reset_index(drop=True)
            fecha_min = master_df[fecha_col].iloc[0].strftime('%d/%m/%Y') if len(master_df) > 0 else '-'
            fecha_max = master_df[fecha_col].iloc[-1].strftime('%d/%m/%Y') if len(master_df) > 0 else '-'
            master_df['Fecha'] = master_df[fecha_col].dt.strftime('%d/%m/%Y')
            self._log(f"Rango de fechas procesado: {fecha_min} - {fecha_max}")

        # 5. Extraer concepto y beneficiario
        desc_col = next((c for c in master_df.columns if 'descrip' in c.lower()), None)
        desc = master_df[desc_col].fillna('').astype(str) if desc_col else pd.Series([''] * len(master_df))
        cb = desc.apply(self.parse_descripcion)
        master_df['Concepto'] = [x[0] for x in cb]
        master_df['Beneficiario'] = [x[1] for x in cb]

        self.history = HistoryManager() if not hasattr(self, 'history') else self.history
        
        # 6. Filtrado de Duplicados Historicos (Hashing)
        total_inicial = len(master_df)
        master_df['record_hash'] = master_df.apply(self.history.generate_hash, axis=1)
        
        # Identificar los que ya estan en DB
        mask_dup = master_df['record_hash'].apply(self.history.is_processed)
        dupes_hist = mask_dup.sum()
        
        if dupes_hist > 0:
            messagebox.showwarning("Registros Duplicados Detectados", 
                                   f"¡Atención!\nSe detectaron {dupes_hist} registros que ya han sido procesados anteriormente.\n\n"
                                   "Estos serán omitidos para evitar duplicidad de recibos.")
            master_df = master_df[~mask_dup].reset_index(drop=True)
            
        self._log(f"Registros totales detectados: {total_inicial}")
        self._log(f"Duplicados historicos omitidos: {dupes_hist}")
        self._log(f"Nuevos registros a procesar: {len(master_df)}")

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
        try:
            df = self.procesar_datos_masivos()
            if df is None or df.empty:
                messagebox.showinfo("Información", "No se encontraron datos nuevos para procesar.")
                return
            self.lbl_status_masivo.configure(text="Procesando... (Por favor espere)", text_color="gray")
            self.btn_generar_masivo.configure(state="disabled")
            self.update()
            threading.Thread(target=self._thread_generar_masivo, args=(df, tpl), daemon=True).start()
        except Exception as e:
            self._log(f"CRITICAL ERROR: {e}")
            messagebox.showerror("Error Critico", f"Falla en el motor de procesamiento: {e}")
            self.btn_generar_masivo.configure(state="normal")

    def _thread_generar_masivo(self, df, tpl):
        self._log("Iniciando motor de automatización...")
        pythoncom.CoInitialize()
        try:
            if not os.path.exists(TARGET_PATH):
                self._log("Creando libro maestro en escritorio...")
                import shutil
                shutil.copy(tpl, TARGET_PATH)
            
            self._log("Escaneando libro maestro para consecutivo...")
            next_block_idx, max_recibo = self.get_master_info()
            self._log(f"Siguiente bloque: {next_block_idx}, Ultimo recibo: {max_recibo}")
            
            self._log("Conectando con Microsoft Excel...")
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
                if (index + 1) % 5 == 0:
                    self._log(f"Procesados {index + 1} de {len(df)}...")
            wb.Save()
            wb.Close(False)
            
            self._log("Guardando registros en historial persistente...")
            # Guardar hashes en el historial despues de proceso exitoso
            processed_hashes = df['record_hash'].tolist()
            self.history.add_records(processed_hashes)
            self._log("✅ Historial actualizado.")

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
