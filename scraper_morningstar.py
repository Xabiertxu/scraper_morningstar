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
        browser = p.chromium.launch(headless=True)
        # Configuración de pantalla para asegurar que el diseño no se rompa
        context = browser.new_context(viewport={"width": 1280, "height": 800})
        page = context.new_page()
        page.set_extra_http_headers({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"})
        
        print(f"[{datetime.now()}] Cargando URL de Morningstar...")
        page.goto(URL_OBJETIVO, timeout=60000)
        
        # --- GESTIÓN DE COOKIES ---
        try:
            # Esperamos brevemente por si salta el banner de cookies
            boton_cookies = page.locator("button:has-text('Aceptar'), button:has-text('Aceptar todo'), #onetrust-accept-btn-handler")
            if boton_cookies.is_visible(timeout=5000):
                boton_cookies.click()
                print("Banner de cookies detectado y aceptado.")
                time.sleep(2)
        except Exception:
            print("No se detectó banner de cookies obligatorio.")

        print("Esperando la carga de los datos de la tabla...")
        # Forzamos una espera genérica para dar margen a la API interna de la web
        page.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(5)
        
        html_content = page.content()
        
        try:
            lista_tablas = pd.read_html(html_content)
            print(f"Tablas encontradas en la página: {len(lista_tablas)}")
            
            # Filtramos para quedarnos con la tabla que realmente contiene los ETFs
            df_etfs = None
            for tabla in lista_tablas:
                if len(tabla) > 1 and ("Name" in tabla.columns or "Nombre" in str(tabla.columns) or "Ticker" in str(tabla.columns)):
                    df_etfs = tabla
                    break
            
            if df_etfs is None and lista_tablas:
                df_etfs = lista_tablas[0] # Por defecto, la primera si tiene filas
                
        except Exception as e:
            print(f"Error al procesar el HTML con Pandas: {e}")
            browser.close()
            exit(1)
        
        if df_etfs is None or df_etfs.empty:
            print("Error: No se pudieron extraer datos estructurados de la tabla.")
            browser.close()
            exit(1)

        print(f"Se han extraído con éxito {len(df_etfs)} filas de ETFs.")
        
        # Consolidar datos
        if os.path.exists(ARCHIVO_SALIDA):
            try:
                df_existente = pd.read_excel(ARCHIVO_SALIDA)
                df_final = pd.concat([df_existente, df_etfs], ignore_index=True).drop_duplicates()
            except Exception:
                df_final = df_etfs
        else:
            df_final = df_etfs
            
        df_final.to_excel(ARCHIVO_SALIDA, index=False)
        print(f"¡Éxito! Archivo '{ARCHIVO_SALIDA}' guardado correctamente con {len(df_final)} registros.")
        browser.close()

if __name__ == "__main__":
    try:
        extraer_datos_morningstar()
    except Exception as e:
        print(f"Fallo crítico en el flujo de scraping: {e}")
        exit(1)
