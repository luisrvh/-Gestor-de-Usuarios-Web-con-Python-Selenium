import tkinter as tk
from tkinter import messagebox
import numpy as np
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

# Ruta del driver Edge (puedes cambiarlo a Chrome o Firefox si lo prefieres)
script_dir = os.path.dirname(os.path.abspath(__file__))
edge_driver_path = os.path.join(script_dir, 'msedgedriver.exe')
options = Options()
options.use_chromium = True
service = Service(executable_path=edge_driver_path)

# Cargar Excel con usuarios
df = pd.read_excel("usuarios.xlsx")

def actualizar_excel(usuario, contrasena):
    df.loc[df['Username'] == usuario, 'Contraseña'] = contrasena
    df.to_excel("usuarios_actualizados.xlsx", index=False)

# Función de automatización principal (Versión 1)
def ejecutar_accion_version1(accion, contrasena_base, incremento_contrasena, label_contador):
    contador = 0
    driver = webdriver.Edge(service=service, options=options)
    driver.get('https://tusitio.com/tu-pagina')  # ← Cambia por la URL de tu plataforma

    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "selectorValue")))

        for index, row in df.iterrows():
            usuario = row['Username']
            contrasena = contrasena_base
            if incremento_contrasena:
                contrasena = contrasena_base + str(index+1)

            if isinstance(usuario, np.int64):
                usuario = str(usuario)

            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "selectorValue"))).clear()
            driver.find_element(By.ID, "selectorValue").send_keys(usuario)

            search_button = driver.find_element(By.XPATH, "//a[@href='javascript: searchFor()']")
            search_button.click()

            time.sleep(10)  # Esperar carga resultados

            # Intentar realizar cambios
            try:
                contador += 1
                label_contador.config(text=f"Usuarios Procesados: {contador}")
                label_contador.update()

                WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, f"ent-password-{contador}"))
                ).clear()

                driver.find_element(By.ID, f"ent-password-{contador}").send_keys(contrasena)

                reset_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, f'//*[@id="ent-content-section-{contador}"]/div[1]/div[2]/div/button[2]'))
                )
                reset_button.click()
                ActionChains(driver).double_click(reset_button).perform()

                time.sleep(10)

                actualizar_excel(usuario, contrasena)

            except Exception as e:
                messagebox.showerror("Error", f"Error procesando usuario {usuario}: {str(e)}")

        messagebox.showinfo("Completado", f"Acción '{accion}' completada con éxito.")

    except Exception as e:
        messagebox.showerror("Error general", str(e))

    finally:
        driver.quit()

# Función secundaria similar (Versión 2)
def ejecutar_accion_version2(accion, contrasena_base, incremento_contrasena, label_contador):
    # Mismo patrón que ejecutar_accion_version1, puedes personalizarlo de igual forma.
    pass  # ← puedes copiar el cuerpo si deseas mantener ambas versiones

def ejecutar_en_hilo(func, accion, contrasena_base, incremento_contrasena, label_contador):
    threading.Thread(target=func, args=(accion, contrasena_base, incremento_contrasena, label_contador)).start()

def ventana_principal(version):
    root = tk.Tk()
    root.title(f"Gestor de Usuarios - Versión {version}")

    tk.Label(root, text="Seleccione una acción:", font=("Arial", 12)).pack(pady=10)

    label_contador = tk.Label(root, text="Usuarios Procesados: 0", font=("Arial", 12))
    label_contador.pack(pady=10)

    tk.Label(root, text="Contraseña base:", font=("Arial", 12)).pack(pady=5)
    entry_contrasena_base = tk.Entry(root, font=("Arial", 12))
    entry_contrasena_base.pack(pady=5)

    var_incremento = tk.IntVar(value=0)
    tk.Radiobutton(root, text="Misma contraseña", variable=var_incremento, value=0).pack(pady=5)
    tk.Radiobutton(root, text="Contraseña incremental", variable=var_incremento, value=1).pack(pady=5)

    if version == 1:
        boton_accion = tk.Button(root, text="Ejecutar",
                                 command=lambda: ejecutar_en_hilo(ejecutar_accion_version1, "cambiar_contraseña", entry_contrasena_base.get(), var_incremento.get(), label_contador),
                                 width=30)
    else:
        boton_accion = tk.Button(root, text="Ejecutar",
                                 command=lambda: ejecutar_en_hilo(ejecutar_accion_version2, "cambiar_contraseña", entry_contrasena_base.get(), var_incremento.get(), label_contador),
                                 width=30)

    boton_accion.pack(pady=5)
    tk.Button(root, text="Salir", command=root.quit, width=30, bg="red", fg="white").pack(pady=10)

    root.mainloop()

def ventana_seleccion_version():
    ventana = tk.Tk()
    ventana.title("Seleccionar Versión")

    tk.Label(ventana, text="Seleccione la versión a usar:", font=("Arial", 14)).pack(pady=15)

    def seleccionar_version(version):
        ventana.destroy()
        ventana_principal(version)

    tk.Button(ventana, text="Versión 1", width=30, height=2, command=lambda: seleccionar_version(1)).pack(pady=10)
    tk.Button(ventana, text="Versión 2", width=30, height=2, command=lambda: seleccionar_version(2)).pack(pady=10)

    ventana.mainloop()

if __name__ == "__main__":
    ventana_seleccion_version()
