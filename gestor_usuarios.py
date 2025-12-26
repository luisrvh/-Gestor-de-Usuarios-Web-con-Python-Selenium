import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import threading
import os

# ---------------- CONFIG ----------------
THEME = "darkly"
APP_TITLE = "Gestión de Usuarios - IT"

# -------- DRIVER --------
script_dir = os.path.dirname(os.path.abspath(__file__))
edge_driver_path = os.path.join(script_dir, 'msedgedriver.exe')
options = Options()
options.use_chromium = True
service = Service(executable_path=edge_driver_path)

# -------- LEER EXCEL --------
df = pd.read_excel("usuarios.xlsx")

# -------- ACTUALIZAR EXCEL --------
def actualizar_excel(usuario, contrasena):
    df.loc[df['Username'] == usuario, 'Contraseña'] = contrasena
    df.to_excel("usuarios_actualizados.xlsx", index=False)

# -------- CERRAR POPUPS --------
def cerrar_ventanas(driver):
    for xp in ["Cerrar", "Close", "Aceptar", "OK"]:
        try:
            driver.find_element(By.XPATH, f"//button[contains(text(), '{xp}')]").click()
        except:
            pass

# -------- LÓGICA SELENIUM --------
def ejecutar_accion(version, contrasena_base, incremento, ui):
    contador = 0
    total = len(df)

    driver = webdriver.Edge(service=service, options=options)
    driver.get('Tu URL')

    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "selectorValue"))
        )

        for index, row in df.iterrows():
            usuario = str(row["Username"])
            contrasena = contrasena_base + (str(index + 1) if incremento else "")

            ui.log(f"Procesando usuario: {usuario}")

            campo = driver.find_element(By.ID, "selectorValue")
            campo.clear()
            campo.send_keys(usuario)
            driver.find_element(By.XPATH, "//a[@href='javascript: searchFor()']").click()
            time.sleep(6)

            if driver.find_elements(By.XPATH, "//*[contains(text(), 'No se encontró')]"):
                ui.log("Usuario no encontrado")
                cerrar_ventanas(driver)
                continue

            contador += 1
            ui.update_progress(contador, total)

            try:
                pass_field = driver.find_element(By.ID, f"ent-password-{contador}")
                pass_field.clear()
                pass_field.send_keys(contrasena)

                btn = driver.find_element(
                    By.XPATH,
                    f'//*[@id="ent-content-section-{contador}"]/div[1]/div[2]/div/button[2]'
                )
                btn.click()
                ActionChains(driver).double_click(btn).perform()

                actualizar_excel(usuario, contrasena)
                ui.log("Contraseña actualizada")

            except:
                ui.log("Error al actualizar usuario")

            cerrar_ventanas(driver)
            time.sleep(2)

        messagebox.showinfo("Proceso Finalizado", "Proceso completado")

    except Exception as e:
        messagebox.showerror("ERROR", str(e))

    finally:
        driver.quit()
        ui.finish()

# -------- APP --------
class App:
    def __init__(self):
        self.app = tb.Window(
            title=APP_TITLE,
            themename=THEME,
            size=(900, 520),
            resizable=(False, False)
        )
        self.container = tb.Frame(self.app)
        self.container.pack(expand=True, fill=BOTH)
        self.mostrar_seleccion()

    def limpiar(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    # -------- PANTALLA 1 --------
    def mostrar_seleccion(self):
        self.limpiar()
        frame = tb.Frame(self.container, padding=30)
        frame.pack(expand=True)

        tb.Label(frame, text=APP_TITLE, font=("Segoe UI", 16, "bold")).pack(pady=10)
        tb.Label(frame, text="Seleccione la versión").pack(pady=10)

        tb.Button(
            frame,
            text="Versión 1 – MES",
            bootstyle=PRIMARY,
            width=30,
            command=lambda: self.mostrar_dashboard(1)
        ).pack(pady=10)

        tb.Button(
            frame,
            text="Versión 2 – Temporal",
            bootstyle=INFO,
            width=30,
            command=lambda: self.mostrar_dashboard(2)
        ).pack(pady=10)

    # -------- PANTALLA 2 --------
    def mostrar_dashboard(self, version):
        self.limpiar()

        main = tb.Frame(self.container)
        main.pack(expand=True, fill=BOTH)

        sidebar = tb.Frame(main, width=240, bootstyle=DARK)
        sidebar.pack(side=LEFT, fill=Y)
        sidebar.pack_propagate(False)

        content = tb.Frame(main, padding=20)
        content.pack(expand=True, fill=BOTH)

        # Sidebar
        sb = tb.Frame(sidebar, padding=15)
        sb.pack(expand=True, fill=BOTH)

        tb.Label(sb, text="Configuración", font=("Segoe UI", 13, "bold")).pack(pady=15)
        tb.Label(sb, text="Contraseña base").pack(anchor=W)

        self.entry_pass = tb.Entry(sb)
        self.entry_pass.pack(fill=X, pady=10)

        self.incremento = tb.IntVar()
        tb.Checkbutton(sb, text="Incrementar por usuario", variable=self.incremento).pack(anchor=W)

        # BOTÓN INICIAR (AQUÍ ESTÁ)
        tb.Button(
            sidebar,
            text="▶ INICIAR PROCESO",
            bootstyle=SUCCESS,
            width=25,
            command=lambda: self.start(version)
        ).pack(side=BOTTOM, pady=20)

        # Content
        tb.Label(content, text="Progreso", font=("Segoe UI", 14, "bold")).pack(anchor=W)
        self.progress = tb.Progressbar(content, maximum=len(df))
        self.progress.pack(fill=X, pady=10)

        self.logbox = tb.Text(content, height=15)
        self.logbox.pack(fill=BOTH, pady=10)

    def log(self, msg):
        self.logbox.insert(END, f"• {msg}\n")
        self.logbox.see(END)

    def update_progress(self, v, t):
        self.progress["value"] = v

    def start(self, version):
        if not self.entry_pass.get():
            messagebox.showwarning("Validación", "Ingrese contraseña")
            return

        threading.Thread(
            target=ejecutar_accion,
            args=(version, self.entry_pass.get(), self.incremento.get(), self),
            daemon=True
        ).start()

    def finish(self):
        pass

    def run(self):
        self.app.mainloop()

# -------- START --------
if __name__ == "__main__":
    App().run()
