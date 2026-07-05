import os
import time
import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# Usamos la URL del buscador moderno que se muestra en tu captura exitosa
URL_BUSCADOR = "https://global.morningstar.com/es/herramientas/buscador/etfs/6555a55e-c0e0-466c-b24d-d9959030e1ce"
ARCHIVO_SALIDA = "lista_completa_etfs_morningstar.xlsx"

def scraping_buscador_etfs():
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Iniciando navegador Playwright (Camuflado)...")
    datos_etfs = []
    
    with sync_playwright() as p:
        # Modo headless con evasión básica de firmas automatizadas
        browser = p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            locale="es-ES",
            viewport={"width": 1440, "height": 900}
        )
        page = context.new_page()
        
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Conectando al buscador moderno de Morningstar...")
        try:
            # Navegar esperando que cargue el árbol DOM inicial
            page.goto(URL_BUSCADOR, wait_until="domcontentloaded", timeout=60000)
            
            # Gestión de cookies en el nuevo formato si aparece
            try:
                page.wait_for_selector("#onetrust-accept-btn-handler", timeout=8000)
                page.click("#onetrust-accept-btn-handler")
                print("-> Banner de cookies aceptado.")
            except:
                pass

            print("-> Esperando a que las filas de ETFs aparezcan en la grilla...")
            
            # En la nueva interfaz de Morningstar, las celdas o filas usan contenedores estructurados o roles de tabla.
            # Esperamos a cualquier elemento de fila interactiva estándar o celda de datos de ETF.
            page.wait_for_selector("div[role='row'], tr, .sal-dp-numeric, div[data-test='grid-row']", timeout=45000)
            
            # Hacemos un scroll suave simulado para asegurar que los scripts pinter la tabla completa
            page.evaluate("window.scrollTo(0, document.body.scrollHeight / 2);")
            time.sleep(4)
            
            # Capturamos el HTML de la página procesada
            html_content = page.content()
            soup = BeautifulSoup(html_content, 'html5lib')
            
            # Buscar el contenedor de datos (soporta tanto tablas nativas como flex/grid de la nueva interfaz)
            filas = soup.find_all(["tr", "div"], attrs={"role": "row"})
            
            # Si no detecta filas por rol, buscamos cualquier elemento fila estándar
            if len(filas) <= 1:
                filas = soup.find_all("tr")
                
            print(f"-> Elementos de fila detectados en el DOM: {len(filas)}")
            
            for index, fila in enumerate(filas):
                # Extraer todas las celdas (tanto td tradicionales como celdas div modernas)
                celdas = fila.find_all(["td", "div"], attrs={"role": "gridcell"})
                if not celdas:
                    celdas = fila.find_all("td")
                    
                if not celdas or len(celdas) < 3:
                    continue
                
                # Mapeo posicional dinámico según tu captura visual:
                # Columna 1: Nombre del ETF | Columna 2: ISIN | Columna 3: Último Precio
                nombre = celdas[0].get_text(strip=True)
                isin = celdas[1].get_text(strip=True)
                precio_raw = celdas[2].get_text(strip=True) if len(celdas) > 2 else "0.0"
                
                # Ignorar filas de cabecera que contengan etiquetas de texto fijas
                if "Nombre" in nombre or "ISIN" in isin or not isin:
                    continue
                
                # Detectar divisa basándonos en el símbolo de la celda de precio
                divisa = "EUR" if "€" in precio_raw else ("USD" if "US$" in precio_raw or "USD" in precio_raw else "N/A")
                
                datos_etfs.append({
                    "ID Morningstar": isin,
                    "Nombre del ETF": nombre,
                    "Ticker": "",  # Reservado para transformaciones analíticas posteriores
                    "ISIN": isin,
                    "Último Precio": precio_raw,
                    "Divisa": divisa,
                    "Bolsa de Cotización": "MCE" if divisa == "EUR" else "Otras"
                })
                
        except Exception as e:
            print(f"❌ Fallo durante la extracción del HTML: {str(e)}")
            raise e
        finally:
            browser.close()

    # --- ESCRITURA EN EXCEL ---
    if datos_etfs:
        df_nuevos = pd.DataFrame(datos_etfs)
        print(f"¡Éxito! Se han extraído correctamente {len(df_nuevos)} ETFs de la página actual.")
    else:
        print("⚠️ No se pudieron leer registros reales en esta ejecución. Aplicando Plan B de contingencia...")
        df_nuevos = pd.DataFrame([{
            "ID Morningstar": "F00000XXXX",
            "Nombre del ETF": "Fila de Control - Comprobar selectores en Local",
            "Ticker": "IE00B4L5Y983",
            "ISIN": "IE00B4L5Y983",
            "Último Precio": "0.0",
            "Divisa": "EUR",
            "Bolsa de Cotización": "MCE"
        }])

    # Consolidar datos evitando duplicar registros de ejecuciones previas basándonos en el ISIN
    if os.path.exists(ARCHIVO_SALIDA):
        try:
            df_existente = pd.read_excel(ARCHIVO_SALIDA)
            df_final = pd.concat([df_existente, df_nuevos], ignore_index=True).drop_duplicates(subset=["ISIN"], keep="last")
        except Exception:
            df_final = df_nuevos
    else:
        df_final = df_nuevos

    df_final.to_excel(ARCHIVO_SALIDA, index=False)
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Archivo '{ARCHIVO_SALIDA}' consolidado con éxito.")

if __name__ == "__main__":
    scraping_buscador_etfs()
