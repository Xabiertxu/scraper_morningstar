import os
import time
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

URL_OBJETIVO = "https://global.morningstar.com/es/herramientas/buscador/etfs/6555a55e-c0e0-466c-b24d-d9959030e1ce"
ARCHIVO_SALIDA = "lista_completa_etfs_morningstar.xlsx"

def extraer_datos_morningstar():
    print(f"[{datetime.now()}] Iniciando navegador virtual Playwright...")
    with sync_playwright() as p:
        # Lanzamos un navegador Chromium en modo headless (invisible)
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        print(f"[{datetime.now()}] Cargando URL de Morningstar...")
        page.goto(URL_OBJETIVO, timeout=60000)
        
        # Esperamos a que la tabla principal de datos se renderice en pantalla
        page.wait_for_selector("table", timeout=30000)
        time.sleep(5) # Margen de seguridad para carga de scripts internos
        
        print("Extrayendo filas de la tabla de ETFs...")
        # Localizamos las tablas y procesamos el HTML visible
        tablas = page.locator("table").all()
        
        if not tablas:
            print("Error: No se encontraron tablas de datos en la página.")
            browser.close()
            return
            
        # Parseo automático de las estructuras de las tablas con Pandas
        html_content = page.content()
        lista_tablas = pd.read_html(html_content)
        
        # Consolidamos las tablas detectadas en un único DataFrame
        df_etfs = lista_tablas[0] if lista_tablas else pd.DataFrame()
        
        if df_etfs.empty:
            print("No se pudieron parsear filas estructuradas.")
            browser.close()
            return

        print(f"Se han extraído con éxito {len(df_etfs)} ETFs de la sesión actual.")
        
        # Guardar o consolidar con histórico previo
        if os.path.exists(ARCHIVO_SALIDA):
            try:
                df_existente = pd.read_excel(ARCHIVO_SALIDA)
                df_final = pd.concat([df_existente, df_etfs], ignore_index=True).drop_duplicates()
            except Exception:
                df_final = df_etfs
        else:
            df_final = df_etfs
            
        df_final.to_excel(ARCHIVO_SALIDA, index=False)
        print(f"¡Éxito! Archivo '{ARCHIVO_SALIDA}' guardado correctamente.")
        browser.close()

if __name__ == "__main__":
    try:
        extraer_datos_morningstar()
    except Exception as e:
        print(f"Fallo crítico en el flujo de scraping: {e}")
        exit(1)
