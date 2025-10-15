# -*- coding: utf-8 -*-
import os
import re
import time
import threading
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

# ---------------------------
# Utilidades
# ---------------------------

def normalize_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip()
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    s = s.upper()
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def candidate_is_phone_col(series: pd.Series) -> bool:
    cnt = 0
    checked = 0
    for v in series.dropna().head(40):
        checked += 1
        s = re.sub(r"\D", "", str(v))
        if 6 <= len(s) <= 13:
            cnt += 1
    if checked == 0:
        return False
    return (cnt / checked) >= 0.5

def format_peru_phone(raw: str) -> str:
    tel = re.sub(r"\D", "", str(raw))
    if tel.startswith("51") and len(tel) == 11:
        return tel
    if len(tel) == 9:
        return "51" + tel
    if len(tel) == 8:
        return "51" + tel
    return tel

def extract_firstname_lastname_from_pdf(filename: str):
    base = os.path.splitext(os.path.basename(filename))[0]
    norm = normalize_text(base)
    parts = [p for p in norm.split(" ") if p]
    if len(parts) >= 2:
        return parts[0], parts[1]
    elif len(parts) == 1:
        return parts[0], ""
    else:
        return "", ""

def find_header_mapping_from_df(df: pd.DataFrame):
    mapping = {"NOMBRES": None, "APELLIDOS": None, "TELEFONO": None}
    col_norm = {c: normalize_text(c) for c in df.columns}
    for c, nc in col_norm.items():
        if ("NOMB" in nc or "NAME" in nc or "NOMBRE" in nc) and mapping["NOMBRES"] is None:
            mapping["NOMBRES"] = c
        if ("APELL" in nc or "LAST" in nc or "APELLIDO" in nc) and mapping["APELLIDOS"] is None:
            mapping["APELLIDOS"] = c
        if any(k in nc for k in ("TEL", "CEL", "NUM", "PHONE")) and mapping["TELEFONO"] is None:
            mapping["TELEFONO"] = c

    if mapping["TELEFONO"] is None:
        for c in df.columns:
            try:
                if candidate_is_phone_col(df[c]):
                    mapping["TELEFONO"] = c
                    break
            except Exception:
                continue

    return mapping

def read_excel_flexible(path):
    try:
        df = pd.read_excel(path, engine="openpyxl")
    except Exception:
        df = pd.read_excel(path)
    mapping = find_header_mapping_from_df(df)
    if all(mapping.values()):
        return df, mapping

    raw = pd.read_excel(path, header=None, engine="openpyxl")
    header_row_index = None
    for r in range(min(8, len(raw))):
        row = raw.iloc[r].fillna("").astype(str).tolist()
        tmp_cols = {}
        for i, cell in enumerate(row):
            nc = normalize_text(cell)
            tmp_cols[i] = nc
        has_name = any("NOMB" in v or "NOMBRE" in v or "NAME" in v for v in tmp_cols.values())
        has_ap = any("APELL" in v or "APELLIDO" in v or "LAST" in v for v in tmp_cols.values())
        has_tel = any("TEL" in v or "CEL" in v or "NUM" in v or "PHONE" in v for v in tmp_cols.values())
        if (has_name and has_ap) or has_tel:
            header_row_index = r
            break
    if header_row_index is not None:
        df2 = pd.read_excel(path, header=header_row_index, engine="openpyxl")
        mapping = find_header_mapping_from_df(df2)
        return df2, mapping

    return df, mapping

# ---------------------------
# App GUI + Lógica
# ---------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Envío de Certificados por WhatsApp (v8)")
        self.geometry("980x720")
        self.minsize(960, 680)
        self.big_font = font.Font(size=14)
        self.small_font = font.Font(size=11)
        self.title_font = font.Font(size=18, weight="bold")

        self.excel_path = None
        self.pdf_folder = None
        self.df_excel = None
        self.mapping = None
        self.participants = []
        self.driver = None
        self.stop_sending = False

        self.build_ui()

    def build_ui(self):
        tk.Label(self, text="Envío de Certificados por WhatsApp", font=self.title_font).pack(pady=10)

        top = tk.Frame(self)
        top.pack(fill="x", pady=6)
        tk.Button(top, text="📁 Seleccionar Excel", font=self.big_font, bg="#4CAF50", fg="white",
                  command=self.select_excel).grid(row=0, column=0, padx=8)
        tk.Button(top, text="📂 Seleccionar carpeta de PDFs", font=self.big_font, bg="#2196F3", fg="white",
                  command=self.select_pdf_folder).grid(row=0, column=1, padx=8)
        tk.Button(top, text="🌐 Abrir WhatsApp Web", font=self.big_font, bg="#128C7E", fg="white",
                  command=self.open_whatsapp).grid(row=0, column=2, padx=8)

        self.lbl_excel = tk.Label(self, text="Archivo Excel: —", font=self.small_font)
        self.lbl_excel.pack(anchor="w", padx=12)
        self.lbl_pdfs = tk.Label(self, text="Carpeta PDFs: —", font=self.small_font)
        self.lbl_pdfs.pack(anchor="w", padx=12)

        frame_list = tk.LabelFrame(self, text="Participantes (marque a quién enviar)", font=self.big_font)
        frame_list.pack(fill="both", expand=True, padx=10, pady=10)

        ctrl = tk.Frame(frame_list)
        ctrl.pack(fill="x")
        self.var_all = tk.BooleanVar()
        tk.Checkbutton(ctrl, text="Seleccionar todo", variable=self.var_all, command=self.toggle_all,
                       font=self.small_font).pack(side="left", padx=8)
        tk.Button(ctrl, text="🔄 Actualizar", font=self.small_font, command=self.combine_and_refresh).pack(side="left", padx=8)

        canvas = tk.Canvas(frame_list)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(frame_list, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)
        self.rows_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")
        self.rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        bottom = tk.Frame(self)
        bottom.pack(fill="x", pady=6)
        self.btn_send = tk.Button(bottom, text="📤 Iniciar envío", font=self.big_font, bg="#f57c00",
                                  fg="white", command=self.start_sending, state="disabled")
        self.btn_send.pack(side="left", padx=8)
        self.btn_stop = tk.Button(bottom, text="⛔ Detener", font=self.big_font, bg="#b00020",
                                  fg="white", command=self.stop_sending_action, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(bottom, orient="horizontal", mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        log_frame = tk.LabelFrame(self, text="Registro", font=self.small_font)
        log_frame.pack(fill="both", expand=False, padx=10, pady=8)
        self.log_box = tk.Text(log_frame, height=8, state="disabled", font=("Arial", 11))
        self.log_box.pack(fill="both", expand=True)

    def log(self, msg: str):
        self.log_box.config(state="normal")
        self.log_box.insert("end", f"{time.strftime('%H:%M:%S')} - {msg}\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    # ------------------ Excel & PDFs ------------------

    def select_excel(self):
        path = filedialog.askopenfilename(title="Selecciona archivo Excel (.xlsx/.xls)", filetypes=[("Excel", "*.xlsx *.xls")])
        if not path:
            return
        self.excel_path = path
        self.lbl_excel.config(text=f"Archivo Excel: {os.path.basename(path)}")
        try:
            df, mapping = read_excel_flexible(path)
            self.df_excel = df
            self.mapping = mapping
            self.log(f"Excel cargado: {path} (filas: {len(df)})")
            mm = {k: (v if v else "NO_DETECTADA") for k, v in mapping.items()}
            self.log(f"Mapeo detectado: {mm}")
            if not mapping.get("TELEFONO"):
                self.log("ATENCIÓN: No se detectó columna de teléfono por nombre. El programa intentará detectar por contenido.")
            if self.pdf_folder:
                self.combine_and_refresh()
        except Exception as e:
            messagebox.showerror("Error lectura Excel", f"No se pudo leer el Excel:\n{e}")
            self.log(f"Error leyendo Excel: {e}")

    def select_pdf_folder(self):
        folder = filedialog.askdirectory(title="Selecciona carpeta donde están los certificados (PDF)")
        if not folder:
            return
        self.pdf_folder = folder
        self.lbl_pdfs.config(text=f"Carpeta PDFs: {folder}")
        self.log(f"Carpeta PDFs seleccionada: {folder}")
        if self.df_excel is not None:
            self.combine_and_refresh()

    def combine_and_refresh(self):
        if self.df_excel is None:
            messagebox.showwarning("Falta Excel", "Primero selecciona el archivo Excel.")
            return
        if self.pdf_folder is None:
            messagebox.showwarning("Falta carpeta PDFs", "Primero selecciona la carpeta de PDFs.")
            return

        mapping = self.mapping or {}
        df = self.df_excel.copy()

        detected = find_header_mapping_from_df(df)
        if not mapping.get("NOMBRES") and detected.get("NOMBRES"):
            mapping["NOMBRES"] = detected["NOMBRES"]
        if not mapping.get("APELLIDOS") and detected.get("APELLIDOS"):
            mapping["APELLIDOS"] = detected["APELLIDOS"]
        if not mapping.get("TELEFONO") and detected.get("TELEFONO"):
            mapping["TELEFONO"] = detected["TELEFONO"]

        possible_names = [c for c in df.columns if any(tok in normalize_text(c) for tok in ("NOMB","NAME"))]
        possible_ap = [c for c in df.columns if any(tok in normalize_text(c) for tok in ("APELL","LAST"))]
        possible_tel = [c for c in df.columns if any(tok in normalize_text(c) for tok in ("TEL","CEL","NUM","PHONE"))]

        if not mapping.get("NOMBRES") and possible_names:
            mapping["NOMBRES"] = possible_names[0]
        if not mapping.get("APELLIDOS") and possible_ap:
            mapping["APELLIDOS"] = possible_ap[0]
        if not mapping.get("TELEFONO") and possible_tel:
            mapping["TELEFONO"] = possible_tel[0]

        if not mapping.get("TELEFONO"):
            for c in df.columns:
                try:
                    if candidate_is_phone_col(df[c]):
                        mapping["TELEFONO"] = c
                        break
                except Exception:
                    continue

        self.mapping = mapping

        try:
            df["NOMBRES_CL"] = df[mapping["NOMBRES"]].apply(normalize_text) if mapping.get("NOMBRES") else ""
            df["APELLIDOS_CL"] = df[mapping["APELLIDOS"]].apply(normalize_text) if mapping.get("APELLIDOS") else ""
            df["CLAVE"] = (df["NOMBRES_CL"].fillna("") + " " + df["APELLIDOS_CL"].fillna("")).str.strip()
            if mapping.get("TELEFONO"):
                df["NUM_WA"] = df[mapping["TELEFONO"]].apply(format_peru_phone)
            else:
                df["NUM_WA"] = ""
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron normalizar columnas:\n{e}")
            self.log(f"Error normalizando: {e}")
            return

        pdf_files = [f for f in os.listdir(self.pdf_folder) if f.lower().endswith(".pdf")]
        pdf_rows = []
        for p in pdf_files:
            fn = os.path.join(self.pdf_folder, p)
            n, a = extract_firstname_lastname_from_pdf(p)
            clave = (normalize_text(n) + " " + normalize_text(a)).strip()
            pdf_rows.append({"PDF": p, "CLAVE": clave, "PATH": fn})

        df_pdfs = pd.DataFrame(pdf_rows)
        if df_pdfs.empty:
            df_pdfs = pd.DataFrame(columns=["PDF","CLAVE","PATH"])

        merged = pd.merge(df, df_pdfs, on="CLAVE", how="left", validate="m:1")
        merged["PDF_PATH"] = merged["PATH"].fillna("")
        merged["ENCONTRADO"] = merged["PDF_PATH"].apply(lambda x: bool(x) and os.path.exists(x))
        merged["NUM_WA"] = merged["NUM_WA"].fillna("")

        self.participants = []
        for _, row in merged.iterrows():
            display_name = (str(row.get(mapping["NOMBRES"], "")) + " " + str(row.get(mapping["APELLIDOS"], ""))).strip()
            self.participants.append({
                "NOMBRES": row.get(mapping["NOMBRES"], ""),
                "APELLIDOS": row.get(mapping["APELLIDOS"], ""),
                "NUM_WA": row.get("NUM_WA", ""),
                "PDF_PATH": row.get("PDF_PATH", ""),
                "ENCONTRADO": bool(row.get("ENCONTRADO", False)),
                "CLAVE": row.get("CLAVE", ""),
                "var": tk.BooleanVar(value=bool(row.get("ENCONTRADO", False))),
                "display": display_name if display_name else row.get("CLAVE", "")
            })

        self.show_participants()
        self.btn_send.config(state="normal")
        self.log(f"Combinación completa. PDFs detectados: {sum(1 for p in self.participants if p['ENCONTRADO'])} / {len(self.participants)}")

    def show_participants(self):
        for w in self.rows_frame.winfo_children():
            w.destroy()
        headers = ["Enviar", "Nombre completo", "Teléfono (WA)", "PDF encontrado"]
        for j, h in enumerate(headers):
            tk.Label(self.rows_frame, text=h, font=self.small_font, borderwidth=0,
                     width=(12 if j==0 else 40)).grid(row=0, column=j, sticky="w", padx=4, pady=2)
        for i, p in enumerate(self.participants, start=1):
            tk.Checkbutton(self.rows_frame, variable=p["var"]).grid(row=i, column=0, padx=6)
            tk.Label(self.rows_frame, text=p["display"], anchor="w").grid(row=i, column=1, sticky="w")
            tk.Label(self.rows_frame, text=p["NUM_WA"]).grid(row=i, column=2)
            tk.Label(self.rows_frame, text=( "✅ " if p["ENCONTRADO"] else "❌ " ) + (os.path.basename(p["PDF_PATH"]) if p["PDF_PATH"] else "No"),
                     anchor="w").grid(row=i, column=3, sticky="w")

    def toggle_all(self):
        val = self.var_all.get()
        for p in self.participants:
            p["var"].set(val)

    # ------------------ WhatsApp & envío ------------------

    def open_whatsapp(self):
        try:
            if self.driver is None:
                options = webdriver.ChromeOptions()
                profile_dir = os.path.join(os.getcwd(), "whatsapp_profile_v8")
                options.add_argument(f"--user-data-dir={profile_dir}")
                options.add_argument("window-size=1200,900")
                options.add_argument("--disable-notifications")
                self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            self.driver.get("https://web.whatsapp.com")
            messagebox.showinfo("WhatsApp Web", "Se abrió WhatsApp Web. Escanea el QR con tu teléfono si es necesario.\n\nCuando aparezcan tus chats, puedes iniciar el envío.")
            def wait_login():
                try:
                    WebDriverWait(self.driver, 180).until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')))
                    self.log("WhatsApp Web: sesión detectada.")
                except Exception:
                    self.log("No se detectó inicio de sesión en el tiempo esperado.")
            threading.Thread(target=wait_login, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error abrir WhatsApp", f"No se pudo abrir WhatsApp Web:\n{e}")
            self.log(f"Error abrir WhatsApp: {e}")

    def start_sending(self):
        selected = [p for p in self.participants if p["var"].get()]
        if not selected:
            messagebox.showwarning("Sin selección", "No hay participantes seleccionados para enviar.")
            return
        if not self.driver:
            messagebox.showwarning("WhatsApp no abierto", "Abre WhatsApp Web primero.")
            return
        if not messagebox.askyesno("Confirmar envío", f"Se enviarán {len(selected)} certificados. ¿Continuar?"):
            return
        self.stop_sending = False
        self.btn_send.config(state="disabled")
        self.btn_stop.config(state="normal")
        threading.Thread(target=self.process_sending, args=(selected,), daemon=True).start()

    def stop_sending_action(self):
        self.stop_sending = True
        self.log("Solicitud de detener el envío recibida. Esperando terminar envío en curso...")

    def send_text_message(self, message_text):
        """Envía un mensaje de texto usando múltiples métodos"""
        try:
            # Método 1: JavaScript directo (más confiable)
            js_code = """
            function sendMessage(text) {
                const textbox = document.querySelector('div[contenteditable="true"][data-tab="10"]');
                if (!textbox) return false;
                
                // Limpiar y establecer el texto
                textbox.innerHTML = '';
                textbox.textContent = text;
                
                // Disparar eventos necesarios
                const inputEvent = new InputEvent('input', { bubbles: true, inputType: 'insertText', data: text });
                textbox.dispatchEvent(inputEvent);
                
                // Pequeña espera para que WhatsApp procese
                setTimeout(() => {
                    const sendBtn = document.querySelector('button[data-tab="11"]');
                    if (sendBtn) {
                        sendBtn.click();
                    }
                }, 100);
                
                return true;
            }
            return sendMessage(arguments[0]);
            """
            result = self.driver.execute_script(js_code, message_text)
            if result:
                time.sleep(2)
                return True
        except Exception as e:
            self.log(f"Error en método JS: {e}")
        
        # Método 2: Selenium tradicional con ActionChains
        try:
            textbox = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
            )
            textbox.click()
            time.sleep(0.5)
            
            # Limpiar cualquier texto previo
            textbox.clear()
            
            # Usar ActionChains para escribir
            actions = ActionChains(self.driver)
            actions.move_to_element(textbox)
            actions.click()
            actions.send_keys(message_text)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            
            time.sleep(2)
            return True
        except Exception as e:
            self.log(f"Error en método ActionChains: {e}")
            
        return False

    def send_pdf_attachment(self, pdf_path):
        """Envía un archivo PDF como adjunto"""
        try:
            # Paso 1: Click en el botón de adjuntar
            attach_btn_selectors = [
                '//div[@title="Adjuntar"]',
                '//button[@data-testid="clip"]',
                '//span[@data-icon="clip"]/..',
                '//div[@aria-label="Adjuntar"]'
            ]
            
            attach_btn = None
            for selector in attach_btn_selectors:
                try:
                    attach_btn = self.driver.find_element(By.XPATH, selector)
                    if attach_btn.is_displayed():
                        break
                except:
                    continue
            
            if not attach_btn:
                self.log("No se encontró el botón de adjuntar")
                return False
            
            attach_btn.click()
            time.sleep(1)
            
            # Paso 2: Localizar el input de archivo y subir
            file_input_selectors = [
                '//input[@accept="*"][@type="file"]',
                '//input[@type="file"]'
            ]
            
            file_input = None
            for selector in file_input_selectors:
                try:
                    file_input = self.driver.find_element(By.XPATH, selector)
                    if file_input:
                        break
                except:
                    continue
            
            if not file_input:
                self.log("No se encontró el input de archivo")
                return False
            
            # Enviar la ruta del archivo
            file_input.send_keys(pdf_path)
            self.log(f"Archivo adjuntado, esperando preview...")
            time.sleep(3)
            
            # Paso 3: Esperar el preview y hacer click en enviar
            send_btn_selectors = [
                '//span[@data-icon="send"]/..',
                '//button[@data-testid="send"]',
                '//div[@aria-label="Enviar"]',
                '//span[@data-testid="send"]/..'
            ]
            
            # Intentar múltiples veces encontrar el botón de enviar
            for attempt in range(5):
                for selector in send_btn_selectors:
                    try:
                        send_btn = self.driver.find_element(By.XPATH, selector)
                        if send_btn.is_displayed() and send_btn.is_enabled():
                            send_btn.click()
                            self.log("Click en botón enviar exitoso")
                            time.sleep(2)
                            return True
                    except:
                        continue
                time.sleep(0.5)
            
            # Si no funciona con clicks, intentar con JavaScript
            self.log("Intentando enviar con JavaScript...")
            js_click = """
            const sendButtons = document.querySelectorAll('button[data-testid="send"], span[data-icon="send"]');
            for (let btn of sendButtons) {
                if (btn.offsetParent !== null) {
                    if (btn.tagName === 'SPAN') {
                        btn.parentElement.click();
                    } else {
                        btn.click();
                    }
                    return true;
                }
            }
            return false;
            """
            result = self.driver.execute_script(js_click)
            if result:
                time.sleep(2)
                return True
            
            self.log("No se pudo hacer click en el botón de enviar")
            return False
            
        except Exception as e:
            self.log(f"Error enviando PDF: {e}")
            return False

    def process_sending(self, list_to_send):
        total = len(list_to_send)
        self.progress["maximum"] = total
        self.progress["value"] = 0
        sent, not_sent = [], []
        out_dir = os.path.dirname(self.excel_path) if self.excel_path else os.getcwd()

        for i, p in enumerate(list_to_send, start=1):
            if self.stop_sending:
                self.log("🚫 Envío detenido por el usuario.")
                break

            name_display = p["display"]
            num = p["NUM_WA"]
            pdf = p["PDF_PATH"]
            tel_digits = re.sub(r"\D", "", str(num))

            if not p["ENCONTRADO"] or not pdf or not os.path.exists(pdf):
                self.log(f"❌ {name_display}: PDF no encontrado.")
                not_sent.append({"NOMBRE": p["NOMBRES"], "TELEFONO": num, "ESTADO": "PDF_NO"})
                self.progress["value"] = i
                continue

            try:
                # Abrir chat
                self.log(f"📱 Abriendo chat de {name_display}...")
                url = f"https://web.whatsapp.com/send?phone={tel_digits}&app_absent=0"
                self.driver.get(url)
                
                # Esperar a que cargue el chat
                WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
                )
                time.sleep(2)

                # Enviar mensaje de saludo
                message_text = f"Hola {p['NOMBRES']}, te envío tu certificado. Saludos."
                self.log(f"💬 Enviando mensaje a {name_display}...")
                
                if self.send_text_message(message_text):
                    self.log(f"✅ Mensaje enviado a {name_display}")
                    time.sleep(2)
                else:
                    self.log(f"⚠️ No se pudo enviar el mensaje a {name_display}")

                # Enviar PDF
                self.log(f"📎 Adjuntando PDF para {name_display}...")
                if self.send_pdf_attachment(pdf):
                    self.log(f"✅ PDF enviado exitosamente a {name_display}")
                    sent.append({"NOMBRE": p["NOMBRES"], "TELEFONO": num, "ESTADO": "ENVIADO"})
                else:
                    self.log(f"❌ Error al enviar PDF a {name_display}")
                    not_sent.append({"NOMBRE": p["NOMBRES"], "TELEFONO": num, "ESTADO": "ERROR_PDF"})

            except Exception as e:
                self.log(f"❌ Error con {name_display}: {e}")
                not_sent.append({"NOMBRE": p["NOMBRES"], "TELEFONO": num, "ESTADO": f"ERROR: {str(e)[:50]}"})

            self.progress["value"] = i
            time.sleep(3)  # Pausa entre envíos

        # Guardar resultados
        self.btn_send.config(state="normal")
        self.btn_stop.config(state="disabled")
        
        if sent:
            pd.DataFrame(sent).to_excel(os.path.join(out_dir, "ENVIADOS.xlsx"), index=False)
        if not_sent:
            pd.DataFrame(not_sent).to_excel(os.path.join(out_dir, "NO_ENVIADOS.xlsx"), index=False)
        
        self.log(f"🗂️ Resultados guardados en {out_dir}")
        messagebox.showinfo("Proceso finalizado", f"✅ Enviados: {len(sent)} | ❌ No enviados: {len(not_sent)}")

    def on_close(self):
        if messagebox.askyesno("Salir", "¿Desea cerrar la aplicación?"):
            try:
                if self.driver:
                    self.driver.quit()
            except:
                pass
            self.destroy()

if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()