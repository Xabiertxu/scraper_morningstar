import os
import json
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

URL_OBJETIVO = "https://global.morningstar.com/es/herramientas/buscador/etfs/6555a55e-c0e0-466c-b24d-d9959030e1ce"
ARCHIVO_SALIDA = "lista_completa_etfs_morningstar.xlsx"

def extraer_datos_morningstar():
    print(f"[{datetime.now()}] Iniciando navegador virtual Playwright...")
    
    datos_api = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1280, "height": 800})
        page = context.new_page()
        
        # Esta función captura de forma transparente las respuestas JSON de la API interna
        def capturar_respuesta(response):
            # Filtramos por las URLs típicas del endpoint de datos (solr o api/v2/search)
            if "search" in response.url or "solr" in response.url:
                try:
                    if response.status == 200:
                        json_data = response.json()
                        # Extraemos los documentos devueltos por la API
                        if "results" in json_data:
                            datos_api.extend(json_data["results"])
                        elif "response" in json_data and "docs" in json_data["response"]:
                            datos_api.extend(json_data["response"]["docs"])
                        print(f"-> Capturado bloque de datos desde la API: {response.url[:60]}...")
                except Exception:
                    pass

        # Vinculamos el interceptor antes de navegar
        page.on("response", capturar_respuesta)
        
        print(f"[{datetime.now()}] Cargando URL de Morningstar e interceptando red...")
        page.goto(URL_OBJETIVO, timeout=60000)
        
        # Esperamos a que la red esté completamente inactiva (lo que significa que la API ya respondió)
        page.wait_for_load_state("networkidle", timeout=30000)
        browser.close()

    if not datos_api:
        print("Error: No se logró capturar ninguna respuesta JSON válida de la API de Morningstar.")
        exit(1)

    print(f"Procesando {len(datos_api)} registros capturados...")
    
    # Mapeamos los campos dinámicos del JSON al formato limpio de Excel solicitado
    lista_limpia = []
    for doc in datos_api:
        etf = {
            # --- SECCIÓN 1: GENERAL ---
            "ID Morningstar": doc.get("Id", doc.get("id", "N/A")),
            "Nombre del ETF": doc.get("Name", doc.get("name", "N/A")),
            "Ticker": doc.get("Ticker", doc.get("ticker", "N/A")),
            "ISIN": doc.get("ISIN", doc.get("isin", "N/A")),
            "Último Precio": doc.get("ClosePrice", doc.get("closePrice", "N/A")),
            "Divisa": doc.get("Currency", doc.get("currency", "N/A")),
            "Bolsa de Cotización": doc.get("ExchangeName", doc.get("exchangeName", "N/A")),
            
            # --- SECCIÓN 2: RENTABILIDAD ---
            "Rentabilidad YTD (%)": doc.get("ReturnYTD", "N/A"),
            "Rentabilidad Anualizada 3 Años (%)": doc.get("ReturnM36", "N/A"),
            
            # --- SECCIÓN 3: RIESGO ---
            "Desviación Estándar 3 Años": doc.get("StandardDeviationThreeYear", "N/A"),
            "Ratio de Sharpe 3 Años": doc.get("SharpeRatioThreeYear", "N/A"),
            
            # --- SECCIÓN 4: SOSTENIBILIDAD ---
            "Sostenibilidad (Rating)": doc.get("SustainabilityRating", "N/A"),
            "Puntuación de Carbono": doc.get("CarbonScore", "N/A")
        }
        lista_limpia.append(etf)

    df_nuevos = pd.DataFrame(lista_limpia)

    # Consolidamos los datos en el Excel
    if os.path.exists(ARCHIVO_SALIDA):
        try:
            df_existente = pd.read_excel(ARCHIVO_SALIDA)
            df_final = pd.concat([df_existente, df_nuevos], ignore_index=True).drop_duplicates(subset=["ISIN"], keep="last")
        except Exception:
            df_final = df_nuevos
    else:
        df_final = df_nuevos

    df_final.to_excel(ARCHIVO_SALIDA, index=False)
    print(f"¡Éxito absoluto! Generado archivo '{ARCHIVO_SALIDA}' con {len(df_final)} ETFs estructurados.")

if __name__ == "__main__":
    extraer_datos_morningstar()
