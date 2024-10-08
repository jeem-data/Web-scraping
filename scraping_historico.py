from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import os
import requests

# Función que devuelve el primer y ultimo dia de cada semana de un año dado
def get_first_and_last_days(year):
    first_days = []
    last_days = []
    first_day_of_year = datetime(year, 1, 1)
    start_of_week = first_day_of_year - timedelta(days=first_day_of_year.weekday())
    while start_of_week.year == year or (start_of_week + timedelta(days=6)).year == year:
        first_day = start_of_week.strftime("%d/%m/%Y")
        first_days.append(first_day)
        last_day = (start_of_week + timedelta(days=6)).strftime("%d/%m/%Y")
        last_days.append(last_day)
        start_of_week += timedelta(days=7)
    return first_days, last_days

# Directorio donde se descargaran los documentos (reemplazar con el directorio deseado)
download_dir = ""

# Elegimos las opciones de lanzamiento del navegador
chrome_options = Options()
chrome_options.add_argument("--disable-search-engine-choice-screen")
prefs = {
    "download.default_directory": download_dir, 
    "download.prompt_for_download": False, 
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("start-maximized")
driver = webdriver.Chrome(options=chrome_options)

# URL donde se encuentran los documentos
driver.get("https://sipi.sic.gov.co/sipi/Extra/Default.aspx")

# A lo largo del script hay una serie de time.sleep(). Esto se hace para que la pagina tenga tiempo de cargar antes de
# intentar realizar la siguiente opción. Hay un codigo que se puede implementar que lo que hace es esperar a que suceda 
# cierto evento, por lo que el tiempo de espera seria optimizado. Se podria implantar en el futuro
time.sleep(2)

# Navega por la página hasta el formulario que hay que rellenar para hacer la búsqueda de los PDFs
element2 = driver.find_element(By.ID, "MainContent_lnkOfficios")
driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element2)
time.sleep(1)
search_text = "Ver Resoluciones"  
element = driver.find_element(By.XPATH, f"//*[contains(text(), '{search_text}')]")
element.click()
time.sleep(2)
button_class_name = "ui-dialog-buttonset"  
button = driver.find_element(By.CLASS_NAME, button_class_name)
button.click()
time.sleep(2)

# Hay 4 tipos de documentos a descargar (descritos a continuacion). Los siguientes números indican su índice dentro del desplegable
# Concede sin oposición = posicion index 8
# Concede con oposición = posicion index 5 
# Niega sin oposicion = posicion index 7
# Niega con oposicion = posicion index 3 
rango_opciones_descarga = [3, 5, 7, 8]
# Prefijos que se le añaden a los nombres de los archivos descargados para facilitar su identificación
prefijos_descarga = ['NCO-', 'CCO-', 'NSO-', 'CSO-']

for y in range(len(rango_opciones_descarga)):
    # Itera sobre todos los años desde 2016 (cuando empieza a haber documentos en esta pagina) hasta 2024
    for x in range(2016, 2025):
        # Obtiene el primer y último dia de cada semana del año actual
        first_days, last_days = get_first_and_last_days(x)

        # Itera por cada semana
        for x in range(len(first_days)):
            # Introduce las fechas correspondientes en el formulario
            input_field = driver.find_element(By.ID, "MainContent_ctrlResolutionSearch_txtResolutionDateStart")  
            input_field.clear()
            input_field.send_keys(first_days[x])
            time.sleep(1)
            input_field = driver.find_element(By.ID, "MainContent_ctrlResolutionSearch_txtResolutionDateEnd")  
            input_field.clear()
            input_field.send_keys(last_days[x])
            button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ui-datepicker-close.ui-state-default.ui-priority-primary.ui-corner-all"))
            )
            button.click()
            time.sleep(2)

            # Introduce el tipo de búsqueda de documento que debe de hacer
            select_element = driver.find_element(By.ID, "MainContent_ctrlResolutionSearch_ddlLetterTemplate")  
            select = Select(select_element)
            select.select_by_index(rango_opciones_descarga[y])
            time.sleep(1)

            # Le da al boton de buscar y espera un poco mas de lo normal para que se carguen bien los resultados
            button_class_name = "ui-button-text"  # Replace with the actual class name of the button
            button = driver.find_element(By.CLASS_NAME, button_class_name)
            button.click()
            time.sleep(4)

            # Primero intenta buscar el aviso que sale en caso de que no haya resultados en la búsqueda
            try:

                close_button = driver.find_element(By.LINK_TEXT, 'close')
                close_button.click()

            # Si no existe (es decir, hay resultados), procede a descargarlos
            except:
                # Localiza la tabla e itera por cada una de sus filas
                table = driver.find_element(By.ID, "MainContent_ctrlResolutionSearch_ctrlDocumentList_gvDocuments")
                rows = table.find_elements(By.XPATH, ".//tr[position() > 1]")  
                for row in rows:
                    try:
                        # Encuentra el elemento en la tercera columna (el PDF a descargar)
                        link_element = row.find_element(By.XPATH, ".//td[3]/a")
                        
                        # Obtiene el enlace de descarga del documento
                        pdf_url = link_element.get_attribute("href")
                        
                        # Obtiene el nombre con el que se descargara el archivo de la lista de prefijos y lo combina con la cuarta columna
                        file_name = prefijos_descarga[y] + row.find_element(By.XPATH, ".//td[4]").text.strip()
                        
                        # Reemplaza caracteres que dan problemas en los nombres de las descargas
                        sanitized_file_name = file_name.replace('/', '-').replace('\\', '-') + ".pdf"
                        
                        # Descarga el archivo
                        file_path = os.path.join(download_dir, sanitized_file_name)
                        response = requests.get(pdf_url)
                        with open(file_path, "wb") as pdf_file:
                            pdf_file.write(response.content)
                        
                        print(f"Downloaded: {sanitized_file_name}")
                        
                    except Exception as e:
                        print(f"Error processing row: {e}")
                        continue
                # Se le pone un tiempo de espera a cada descarga dado que si se intentan descargar demasiados archivos a la vez da problemas
                time.sleep(6)

# Antes de cerrar el programa se espera unos segundos a que cualquier descarga en progreso pueda terminar
time.sleep(10)
driver.quit()