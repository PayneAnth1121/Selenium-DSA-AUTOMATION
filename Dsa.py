from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Configurar el controlador de Edge con la ruta proporcionada
service = Service('C:\\Users\\L_Barnett\\OneDrive - Dell Technologies\\Documents\\Edge Webdriver\\msedgedriver.exe')
driver = webdriver.Edge(service=service)

# Abrir la página web
driver.get('https://sales.dell.com/salesux/customer/search?buid=11&country=US&language=EN&region=US')

# Esperar a que la página cargue
time.sleep(5)

# Paso 1: Esperar a que el elemento esté presente y sea interactivo
wait = WebDriverWait(driver, 10)
search_customer = wait.until(EC.element_to_be_clickable((By.ID, 'intuitive_search_input')))
search_customer.clear()
search_customer.send_keys('16985120')
search_customer.send_keys(Keys.RETURN)
# Esperar a que los resultados carguen
time.sleep(25)

customer_number = wait.until(EC.element_to_be_clickable((By.ID, 'searchSolutions_customerNumber')))
customer_number.clear()
customer_number.send_keys('16985120')
customer_number.send_keys(Keys.RETURN)
# Cerrar el navegador
driver.quit()