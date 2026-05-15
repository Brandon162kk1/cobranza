#-- Froms --
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
#-- Imports --
import os
import time
import pandas as pd
import random

#-------------- Lock File -------------
# Ruta del directorio compartido entre contenedores
SYNC_DIR = "/app/sync"
# Asegurar que exista dentro del volumen (solo la primera vez)
os.makedirs(SYNC_DIR, exist_ok=True)
# Archivo de lock compartido
LOCK_FILE = os.path.join(SYNC_DIR, "session.lock")

def mover_y_hacer_click_simple(driver, elemento, steps=6, pause_between=0.06):
    """
    Mueve el mouse en 'steps' pasos hacia el centro del elemento y hace click.
    driver: tu instancia de webdriver
    elemento: WebElement destino
    """
    action = ActionChains(driver)
    # asegurarnos que el elemento esté visible en pantalla
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
    time.sleep(random.uniform(0.15, 0.45))

    # posiciona el mouse sobre el elemento (move_to_element genera mouseover)
    action.move_to_element(elemento).pause(random.uniform(0.05, 0.18)).perform()

    # pequeños movimientos aleatorios alrededor antes del click
    for _ in range(steps):
        offset_x = random.randint(-6, 6)
        offset_y = random.randint(-6, 6)
        action.move_by_offset(offset_x, offset_y).pause(pause_between)
    # volver al elemento y click
    action.move_to_element(elemento).pause(random.uniform(0.08, 0.2)).click().perform()

def escribir_lento(elemento, texto,min_delay=0.7, max_delay=0.9):
    """Envía texto carácter por carácter con retrasos aleatorios."""
    for letra in texto:
        elemento.send_keys(letra)
        time.sleep(random.uniform(min_delay, max_delay))

def validar_pagina(driver):

    asunto = ""

    try:

        # Validar si aparece el mensaje de error en el body
        if "The requested URL was rejected. Please consult with your administrator." in driver.page_source:
            asunto = "❌ Página web de La Positiva fuera de Servicio"
            return False,asunto

        # Validar si aparece el campo de usuario
        user_field = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.NAME, "txtUsuario")))
        if user_field:
            asunto = "Aparecio otra página web"
            return False,asunto

    except TimeoutException:
        return True,asunto

def acquire_lock():
    try:
        # Intentar crear el archivo de lock
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_RDWR)
        os.write(fd, str(os.getpid()).encode())
        os.close(fd)
        return True
    except FileExistsError:
        # Ya existe => alguien más tiene el lock
        return False

def release_lock():
    try:
        os.remove(LOCK_FILE)
        print("🔓 Lock liberado.")
    except FileNotFoundError:
        pass

def wait_for_lock():
    print("🔒 Esperando que se libere el lock...")
    while True:
        if not os.path.exists(LOCK_FILE):
            if acquire_lock():
                print("✅ Lock adquirido.")
                return True
        time.sleep(5)  # espera 5 segundos antes de volver a intentar

def parse_fecha(fecha_raw):
    # Si ya es Timestamp, la retorna igual
    if isinstance(fecha_raw, pd.Timestamp):
        return fecha_raw
    fecha_str = str(fecha_raw)
    # Si la fecha tiene "-" en la posición 4, probablemente es "YYYY-MM-DD"
    if "-" in fecha_str and fecha_str[4] == "-":
        # Formato ISO (YYYY-MM-DD), NO uses dayfirst
        return pd.to_datetime(fecha_str, dayfirst=False)
    else:
        # Formato DD/MM/YYYY o similar, usa dayfirst
        return pd.to_datetime(fecha_str, dayfirst=True)
