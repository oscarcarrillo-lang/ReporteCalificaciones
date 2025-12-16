from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import pandas as pd # Necesario para consolidar datos (pip install pandas odfpy openpyxl)
import glob

# --- CONFIGURACIÓN DE CARPETAS ---

# Detectar carpeta de descargas del usuario y crear carpeta temporal
user_home = os.path.expanduser("~")
download_folder = os.path.join(user_home, "Downloads", "TempReportesCUN")

if not os.path.exists(download_folder):
    os.makedirs(download_folder)
    print(f"Carpeta creada: {download_folder}")
else:
    print(f"Usando carpeta existente: {download_folder}")
    # Opcional: Limpiar carpeta antes de empezar
    # files = glob.glob(os.path.join(download_folder, "*"))
    # for f in files: os.remove(f)

# --- DEFINICIÓN DE FUNCIONES ---

def reportDownload(driver, wait, base_url, report_id):
    """
    Navega a la URL y descarga el archivo en la carpeta configurada.
    """
    try:
        print(f"\n--- Procesando Reporte ID: {report_id} ---")
        target_url = f"{base_url}/grade/export/ods/index.php?id={report_id}"
        
        driver.get(target_url)

        # Buscar botón de descarga
        print("Buscando botón de descarga...")
        boton = wait.until(EC.element_to_be_clickable((By.ID, "id_submitbutton")))
        
        boton.click()
        print(f"Descarga iniciada para ID {report_id}")

        # Esperar a que la descarga termine (forma simple)
        # Una forma más robusta sería monitorear la carpeta hasta que desaparezca el archivo .crdownload
        time.sleep(5) 
        
    except Exception as e:
        print(f"Error al descargar el reporte {report_id}: {e}")

def consolidarArchivos(source_folder):
    """
    Lee todos los archivos .ods/.xlsx de la carpeta y los une en uno solo.
    """
    print("\n" + "="*40)
    print("INICIANDO CONSOLIDACIÓN DE ARCHIVOS")
    print("="*40)

    # Buscar archivos ODS (la URL dice /ods/) o Excel
    archivos = glob.glob(os.path.join(source_folder, "*.ods")) + glob.glob(os.path.join(source_folder, "*.xlsx"))
    
    if not archivos:
        print("No se encontraron archivos para consolidar.")
        return

    lista_dfs = []
    
    for archivo in archivos:
        try:
            print(f"Leyendo: {os.path.basename(archivo)}")
            # Leemos el archivo. Si es ODS requiere engine='odf'
            if archivo.endswith('.ods'):
                df = pd.read_excel(archivo, engine="odf")
            else:
                df = pd.read_excel(archivo)
            
            # Opcional: Agregar una columna para saber de qué archivo vino
            # df['Origen_Archivo'] = os.path.basename(archivo)
            
            lista_dfs.append(df)
        except Exception as e:
            print(f"Error leyendo {archivo}: {e}")

    if lista_dfs:
        try:
            # Unir todos los dataframes
            consolidado = pd.concat(lista_dfs, ignore_index=True)
            
            # Ruta de salida
            output_path = os.path.join(source_folder, "Consolidado_Final.xlsx")
            
            # Guardar en Excel
            consolidado.to_excel(output_path, index=False)
            print(f"\n¡ÉXITO! Archivo consolidado guardado en:\n{output_path}")
            print(f"Total de registros procesados: {len(consolidado)}")
        except Exception as e:
            print(f"Error al guardar el consolidado: {e}")

# --- BLOQUE PRINCIPAL ---

# Configuración del navegador para descargas automáticas
options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_folder, # Define la carpeta de descarga
    "download.prompt_for_download": False,         # No preguntar dónde guardar
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

# Inicializar el driver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 10)

try:
    # ---------------------------------------------------------
    # PASO 1: LOGIN 
    # ---------------------------------------------------------
    base_url = "https://cdigital.cun.edu.co" 
    print("Navegando al login...")
    driver.get(f"{base_url}/login/index.php")
    
    print("Ingresando credenciales...")
    wait.until(EC.element_to_be_clickable((By.ID, "username"))).send_keys("bral_239") 
    driver.find_element(By.ID, "password").send_keys("Star2025*")

    # Enviar formulario
    try:
        driver.find_element(By.ID, "loginbtn").click()
    except:
        from selenium.webdriver.common.keys import Keys
        driver.find_element(By.ID, "password").send_keys(Keys.RETURN)

    time.sleep(3) # Esperar carga del dashboard

    # ---------------------------------------------------------
    # LECTURA DE DATOS Y DESCARGAS
    # ---------------------------------------------------------
    report_ids = []
    if os.path.exists("data.txt"):
        with open("data.txt", "r") as file:
            report_ids = [line.strip() for line in file.readlines() if line.strip()]
        print(f"IDs encontrados: {len(report_ids)}")
    else:
        print("Error: data.txt no existe.")

    # Descargar cada reporte
    for r_id in report_ids:
        reportDownload(driver, wait, base_url, r_id)

    # Esperar un poco a que terminen todas las descargas pendientes antes de consolidar
    time.sleep(5)

    # ---------------------------------------------------------
    # PASO 3: CONSOLIDACIÓN
    # ---------------------------------------------------------
    # Ejecutamos la consolidación
    consolidarArchivos(download_folder)

except Exception as e:
    print(f"Ocurrió un error crítico: {e}")

finally:
    print("\nProceso finalizado. Cerrando navegador...")
    if 'driver' in locals():
        driver.quit()
