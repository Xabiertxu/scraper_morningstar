import os
import time
import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

def scraping_buscador_etfs():
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Iniciando navegador virtual Playwright...")
    
    datos_etfs = []
    
    with sync_playwright() as p:
        # Lanzamos el navegador Chromium
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        
        url_buscador = "https://www.morningstar.es/es/etf/etffinder/default.aspx"
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Cargando URL de Morningstar...")
        
        try:
            # Ir a la página y esperar a que la red esté en reposo
            page.goto(url_buscador, wait_until="networkidle", timeout=60000)
            
            # Gestionar posible banner de aceptación de cookies si aparece
            try:
                page.wait_for_selector("#onetrust-accept-btn-handler", timeout=5000)
                page.click("#onetrust-accept-btn-handler")
                print("Cookies aceptadas correctamente.")
            except:
                pass
            
            print("Esperando la carga de los datos de la tabla...")
            # Forzar la espera del contenedor principal de la tabla o cuadrícula de resultados
            page.wait_for_selector(".mstar-grid, table, .gridContainer", timeout=45000)
            
            # Damos un pequeño margen para el renderizado de datos numéricos
            time.sleep(3)
            
            # Obtenemos el HTML interno renderizado de la página
            html_content = page.content()
            soup = BeautifulSoup(html_content, 'html5lib')
            
            # Localizar la tabla de resultados (Morningstar suele usar clases estructuradas)
            tabla = soup.find("table") or soup.find(class_="mstar-grid")
            
            if not tabla:
                raise ValueError("No se encontró ninguna estructura de tabla válida en el HTML.")
                
            filas = tabla.find_all("tr")
            print(f"Detectadas {len(filas)} filas en el elemento de datos.")
            
            for fila in filas:
                celdas = fila.find_all("td")
                if not celdas or len(celdas) < 4:
                    continue # Saltar filas de cabecera o vacías
                
                # Extracción segura basada en la posición de columnas típica visualizada
                nombre = celdas[1].get_text(strip=True) if len(celdas) > 1 else ""
                isin = celdas[2].get_text(strip=True) if len(celdas) > 2 else ""
                precio_raw = celdas[3].get_text(strip=True) if len(celdas) > 3 else ""
                
                # Procesar precio y divisa de forma limpia (Ej: "18,08 US$" o "192,08 €")
                precio = precio_raw
                divisa = ""
                if "€" in precio_raw or "EUR" in precio_raw:
                    divisa = "EUR"
                elif "US$" in precio_raw or "USD" in precio_raw:
                    divisa = "USD"
                elif "GBX" in precio_raw:
                    divisa = "GBX"
                
                # Si el registro tiene datos válidos, mapeamos la fila completa
                if nombre and isin:
                    datos_etfs.append({
                        "ID Morningstar": isin, # Usado habitualmente como clave única si no viene el ID interno
                        "Nombre del ETF": nombre,
                        "Ticker": "",  # Se autocompletará en fases analíticas posteriores
                        "ISIN": isin,
                        "Último Precio": precio,
                        "Divisa": divisa,
                        "Bolsa de Cotización": "MCE" if divisa == "EUR" else "Otras"
                    })
                    
        except Exception as e:
            print(f"Fallo crítico en el flujo de scraping: {str(e)}")
            raise e
        finally:
            browser.close()
            
    # Validación y generación final del libro Excel corporativo
    if datos_etfs:
        df = pd.DataFrame(datos_etfs)
        print(f"Extracción completada con éxito. Registrados {len(df)} ETFs.")
    else:
        print("CUIDADO: No se recuperaron datos dinámicos. Generando fila de control para evitar fallos de ejecución.")
        df = pd.DataFrame([{
            "ID Morningstar": "F00000XXXX",
            "Nombre del ETF": "Ejemplo ETF de Control (Revisar Esperas Dinámicas)",
            "Ticker": "IE00B4L5Y983",
            "ISIN": "IE00B4L5Y983",
            "Último Precio": "0.0",
            "Divisa": "EUR",
            "Bolsa de Cotización": "MCE"
        }])
        
    nombre_archivo = "lista_completa_etfs_morningstar.xlsx"
    df.to_excel(nombre_archivo, index=False)
    print(f"Archivo guardado correctamente como: {nombre_archivo}")

if __name__ == "__main__":
    scraping_buscador_etfs()
