import time
import os
import msvcrt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, WebDriverException
import logging
import sys
import pandas as pd
import shutil

class DellSalesPortalAutomation:
    def __init__(self, driver_path, downloads_path):
        self.driver_path = driver_path
        self.downloads_path = downloads_path
        self.consolidated_data_path = r"C:\Users\L_Barnett\OneDrive - Dell Technologies\Documents\My coding Projects\DSA AUTOMATION\Data.xlsx"
        self.driver = None
        self.wait = None
        self.main_window = None
        self.setup_logging()

    def setup_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)

    def initialize_driver(self):
        try:
            service = Service(self.driver_path)
            self.driver = webdriver.Edge(service=service)
            self.wait = WebDriverWait(self.driver, 20)
            self.driver.maximize_window()
            self.logger.info("WebDriver initialized successfully and window maximized")
        except WebDriverException as e:
            self.logger.error(f"Failed to initialize WebDriver: {str(e)}")
            sys.exit(1)

    def navigate_to_sales_portal(self):
        try:
            self.driver.get('https://sales.dell.com/salesux/customer/search?buid=11&country=US&language=EN&region=US')
            self.main_window = self.driver.current_window_handle
            self.logger.info("Successfully navigated to Dell Sales Portal and stored main window handle")
            time.sleep(3)  # Give time for the page to fully load
        except WebDriverException as e:
            self.logger.error(f"Failed to navigate to the sales portal: {str(e)}")
            self.close_browser()
            sys.exit(1)

    def enter_customer_number(self, customer_number):
        try:
            selectors = [
                f"#customerId{customer_number}",
                "input.dds__input-text[acceptonlyalphanumeric]",
                "[id^='customerId']"
            ]
            
            input_found = False
            for selector in selectors:
                try:
                    customer_input = self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    
                    self.wait.until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    
                    customer_input.clear()
                    customer_input.send_keys(customer_number)
                    time.sleep(1)
                    
                    self.click_search_button()
                    
                    self.logger.info(f"Customer number entered successfully: {customer_number}")
                    input_found = True
                    break
                    
                except Exception as e:
                    self.logger.warning(f"Failed to input using selector {selector}: {str(e)}")
                    continue
            
            if not input_found:
                raise Exception("Unable to input customer number with any selector")

        except Exception as e:
            self.logger.error(f"Error entering customer number: {str(e)}")
            raise
        
    def click_search_button(self):
        try:
            selectors = [
                "[id^='customerSearch_searchAction']",
                "button[data-analytics-name='customer-search-form-submit']",
                "button.dds__button.dds__mr-3"
            ]
            
            button_found = False
            for selector in selectors:
                try:
                    wait = WebDriverWait(self.driver, 30)
                    search_button = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
                    time.sleep(1)
                    
                    self.driver.execute_script("arguments[0].click();", search_button)
                    
                    self.logger.info(f"Search button clicked successfully using selector: {selector}")
                    button_found = True
                    break
                    
                except Exception as e:
                    self.logger.warning(f"Failed to click using selector {selector}: {str(e)}")
                    continue
            
            if not button_found:
                raise Exception("Unable to click search button with any selector")

        except Exception as e:
            self.logger.error(f"Error clicking search button: {str(e)}")
            raise

    def click_order_list_tab(self):
        try:
            time.sleep(3)  # Increased wait time
            
            selectors = [
                "button[data-analytics-name='customer-details-tab-orderlist']",
                "#tab-1-2",
                "button[role='tab'] span[title='Order List']"
            ]
            
            button_found = False
            for selector in selectors:
                try:
                    wait = WebDriverWait(self.driver, 30)
                    
                    order_list_button = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", order_list_button)
                    time.sleep(1)
                    
                    self.driver.execute_script("arguments[0].click();", order_list_button)
                    
                    self.logger.info(f"Successfully clicked Order List tab using selector: {selector}")
                    button_found = True
                    break
                    
                except Exception as e:
                    self.logger.warning(f"Failed to click using selector {selector}: {str(e)}")
                    continue
            
            if not button_found:
                raise Exception("Unable to click Order List tab with any selector")

        except Exception as e:
            self.logger.error(f"Error clicking Order List tab: {str(e)}")
            raise

    def click_view_120_days(self):
        try:
            time.sleep(3)  # Increased wait time
            
            selectors = [
                "i[data-analytics-name='OrderList-View120days']",
                "button.dds__button--tertiary i.dds__icon--eye-view-on",
                "button.dds__button--tertiary:has(i[data-analytics-name='OrderList-View120days'])"
            ]
            
            button_found = False
            for selector in selectors:
                try:
                    wait = WebDriverWait(self.driver, 30)
                    
                    view_120_button = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", view_120_button)
                    time.sleep(1)
                    
                    # Store the current window handle before clicking
                    original_window = self.driver.current_window_handle
                    
                    self.driver.execute_script("arguments[0].click();", view_120_button)
                    
                    # Wait for the new window/tab to open
                    wait.until(EC.number_of_windows_to_be(2))
                    
                    # Switch to the new window/tab
                    for window_handle in self.driver.window_handles:
                        if window_handle != original_window:
                            self.driver.switch_to.window(window_handle)
                            break
                    
                    self.logger.info(f"Successfully clicked View 120 days button and switched to new tab")
                    button_found = True
                    break
                    
                except Exception as e:
                    self.logger.warning(f"Failed to click using selector {selector}: {str(e)}")
                    continue
            
            if not button_found:
                raise Exception("Unable to click View 120 days button with any selector")

        except Exception as e:
            self.logger.error(f"Error clicking View 120 days button: {str(e)}")
            raise

    def adjust_date_range(self):
        try:
            time.sleep(5)  # Wait for page load
            
            # Wait for the loading spinner to disappear
            try:
                wait = WebDriverWait(self.driver, 30)
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "dds__loading-spinner")))
            except:
                pass

            # Find From Date input and wait for it to be clickable
            from_date_input = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#text-input-from-date"))
            )

            # Get current date
            current_date = time.strftime("%m/%d/%Y")
            
            # Calculate date from one year ago
            current_month, current_day, current_year = map(int, current_date.split('/'))
            one_year_ago = f"{current_month:02d}/{current_day+1}/{current_year-1}"
            
            # Clear the input field using multiple methods to ensure it works
            self.driver.execute_script("arguments[0].value = '';", from_date_input)
            from_date_input.clear()
            actions = webdriver.ActionChains(self.driver)
            actions.move_to_element(from_date_input).click().perform()
            from_date_input.send_keys(Keys.CONTROL + "a")
            from_date_input.send_keys(Keys.DELETE)
            time.sleep(1)

            # Type the date character by character with small delays
            for char in one_year_ago:
                from_date_input.send_keys(char)
                time.sleep(0.1)

            # Press Tab to move focus and trigger validation
            from_date_input.send_keys(Keys.TAB)
            time.sleep(1)
            
            # Verify the input was accepted
            actual_value = from_date_input.get_attribute('value')
            if actual_value != one_year_ago:
                self.logger.warning(f"Date input verification failed. Expected: {one_year_ago}, Got: {actual_value}")
                
                # Retry once if the first attempt failed
                from_date_input.clear()
                actions.move_to_element(from_date_input).click().perform()
                time.sleep(0.5)
                from_date_input.send_keys(one_year_ago)
                from_date_input.send_keys(Keys.TAB)
                
            self.logger.info(f"Set From Date to: {one_year_ago}")

        except Exception as e:
            self.logger.error(f"Error adjusting date range: {str(e)}")
            raise 
    
    def buttonsearch(self):
        try:
            selectors = [
            "button[role='button'][type='submit'][name='btnSearch'][data-analytics-name='Historymfe-OrderSearch-Search']",
            "button.dds__button.dds__button--sm.dds__mr-3.width-100",
            "button[data-analytics-name='Historymfe-OrderSearch-Search']",
            "button[name='btnSearch']"
        ]
            
            button_found = False
            for selector in selectors:
                try:
                    wait = WebDriverWait(self.driver, 30)
                    search_button = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )

                    self.driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
                    time.sleep(1)

                    self.driver.execute_script("arguments[0].click();", search_button)

                    self.logger.info(f"Search button clicked successfully using selector: {selector}")
                    button_found = True
                    break

                except Exception as e:
                    self.logger.warning(f"Failed to click using selector {selector}: {str(e)}")
                    continue

            if not button_found:
                raise Exception("Unable to click search button with any selector")
                
        except Exception as e:
            self.logger.error(f"Error clicking search button: {str(e)}")
            raise

    def click_download_excel_button(self):
        try:
            time.sleep(3)  # Wait for results to load
            selectors = [
            "#exportExcel",
            "button[data-analytics-name='Historymfe-OrderSearch-DownloadExcel']",
            "button.dds__button.dds__button--sm.dds__button--tertiary"
        ]
            
            button_found = False
            for selector in selectors:
                try:
                    wait = WebDriverWait(self.driver, 30)
                    download_button = wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR,selector))
                    )

                    self.driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
                    time.sleep(1)

                    self.driver.execute_script("arguments[0].click();", download_button)

                    self.logger.info(f"Download to Excel button clicked successfully using selector: {selector}")
                    button_found = True
                    
                    # Wait for download to complete
                    time.sleep(5)
                    break
                
                except Exception as e:
                    self.logger.warning(f"Failed to click using selector {selector}: {str(e)}")
                    continue

            if not button_found:
                raise Exception("Unable to click Download to Excel button with any selector")
            
        except Exception as e:
            self.logger.error(f"Error clicking Download to Excel button: {str(e)}")
            raise  # Cambiado para propagar el error y que se detecte en el procesamiento principal

    def process_downloaded_excel(self, customer_id):
        """
        Procesa el archivo Excel descargado, lo consolida en Data.xlsx y lo elimina
        """
        try:
            # Buscar el archivo Excel más reciente en la carpeta de descargas
            files = [f for f in os.listdir(self.downloads_path) if f.endswith('.xlsx')]
            if not files:
                self.logger.warning(f"No se encontró archivo Excel para el cliente {customer_id}")
                return

            latest_file = max([os.path.join(self.downloads_path, f) for f in files], key=os.path.getctime)
            
            # Leer el archivo descargado
            df_new = pd.read_excel(latest_file)
            
            # Agregar columna de customer_id
            df_new['Customer_ID'] = customer_id
            
            # Si el archivo Data.xlsx no existe, crearlo
            if not os.path.exists(self.consolidated_data_path):
                df_new.to_excel(self.consolidated_data_path, index=False)
                self.logger.info(f"Creado nuevo archivo Data.xlsx con datos del cliente {customer_id}")
            else:
                # Leer el archivo existente y concatenar
                df_existing = pd.read_excel(self.consolidated_data_path)
                df_consolidated = pd.concat([df_existing, df_new], ignore_index=True)
                
                # Guardar archivo consolidado
                df_consolidated.to_excel(self.consolidated_data_path, index=False)
                self.logger.info(f"Datos del cliente {customer_id} agregados a Data.xlsx")
            
            # Eliminar archivo descargado
            os.remove(latest_file)
            self.logger.info(f"Archivo descargado para cliente {customer_id} eliminado")

        except Exception as e:
            self.logger.error(f"Error procesando Excel para cliente {customer_id}: {str(e)}")

    def close_browser(self):
        """Cierra completamente el navegador"""
        if self.driver:
            try:
                self.logger.info("Closing browser completely")
                self.driver.quit()
                self.driver = None
                self.wait = None
                self.main_window = None
                time.sleep(2)  # Tiempo para asegurar que el navegador se cierre correctamente
            except Exception as e:
                self.logger.error(f"Error closing browser: {str(e)}")

def process_downloaded_excel(self, customer_id):
    """
    Procesa el archivo Excel descargado, lo consolida en Data.xlsx y lo elimina
    """
    try:
        # Buscar el archivo Excel más reciente en la carpeta de descargas
        files = [f for f in os.listdir(self.downloads_path) if f.endswith('.xlsx')]
        if not files:
            self.logger.warning(f"No se encontró archivo Excel para el cliente {customer_id}")
            return

        latest_file = max([os.path.join(self.downloads_path, f) for f in files], key=os.path.getctime)
        
        # Leer el archivo descargado
        df_new = pd.read_excel(latest_file)
        
        # Validar que el DataFrame no esté vacío
        if df_new.empty:
            self.logger.warning(f"El archivo descargado para el cliente {customer_id} está vacío")
            os.remove(latest_file)
            return
        
        # Agregar columna de customer_id
        df_new['Customer_ID'] = customer_id
        
        # Si el archivo Data.xlsx no existe, crearlo
        if not os.path.exists(self.consolidated_data_path):
            df_new.to_excel(self.consolidated_data_path, index=False)
            self.logger.info(f"Creado nuevo archivo Data.xlsx con datos del cliente {customer_id}")
        else:
            # Leer el archivo existente y concatenar
            df_existing = pd.read_excel(self.consolidated_data_path)
            
            # Eliminar duplicados si es necesario
            df_consolidated = pd.concat([df_existing, df_new], ignore_index=True)
            df_consolidated.drop_duplicates(inplace=True)
            
            # Guardar archivo consolidado
            df_consolidated.to_excel(self.consolidated_data_path, index=False)
            self.logger.info(f"Datos del cliente {customer_id} agregados a Data.xlsx")
        
        # Eliminar archivo descargado
        os.remove(latest_file)
        self.logger.info(f"Archivo descargado para cliente {customer_id} eliminado")

    except Exception as e:
        self.logger.error(f"Error procesando Excel para cliente {customer_id}: {str(e)}")
        # Intentar eliminar el archivo descargado incluso si hay un error
        try:
            os.remove(latest_file)
        except:
            pass

def read_customer_numbers_from_excel(excel_path):
    """
    Lee los números de cliente desde un archivo Excel.
    
    Parámetros:
    excel_path (str): Ruta completa al archivo Excel
    
    Retorna:
    list: Lista de números de cliente como cadenas
    """
    try:
        # Importar pandas aquí para asegurar que esté disponible
        import pandas as pd
        
        # Configurar logging
        import logging
        logger = logging.getLogger(__name__)
        
        # Leer Excel file
        df = pd.read_excel(excel_path)
        
        # Imprimir columnas para debugging
        logger.info(f"Columnas en el Excel: {list(df.columns)}")
        
        # Extraer números de cliente desde la primera columna, comenzando desde la segunda fila
        # Usar iloc para indexación basada en posición
        customer_numbers = df.iloc[1:, 0].tolist()
        
        # Convertir a cadenas, manejar valores NaN
        customer_numbers = [
            str(int(num)) if pd.notna(num) and not pd.isnull(num) 
            else None for num in customer_numbers
        ]
        
        # Filtrar valores nulos
        customer_numbers = [num for num in customer_numbers if num is not None]
        
        # Log de números de cliente encontrados
        logger.info(f"Números de cliente encontrados: {customer_numbers}")
        
        return customer_numbers
    
    except Exception as e:
        # Logging de errores
        import logging
        logger = logging.getLogger(__name__)
        logger.error(f"Error al leer números de cliente desde Excel: {str(e)}")
        
        # Imprimir detalles del error para debugging
        import traceback
        logger.error(traceback.format_exc())
        
        return []
        
def main():
    # Configurar logging básico
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("dell_sales_portal_automation.log"),  # Log a un archivo
            logging.StreamHandler()  # Log a consola
        ]
    )
    
    # Excel file path
    excel_path = r"C:\Users\L_Barnett\OneDrive - Dell Technologies\Documents\My coding Projects\DSA AUTOMATION\Missing contacts.xlsx"
    
    # Edge WebDriver path
    webdriver_path = r'C:\Users\L_Barnett\OneDrive - Dell Technologies\Documents\Edge Webdriver\msedgedriver.exe'
    
    # Downloads path
    downloads_path = r"C:\Users\L_Barnett\Downloads"
    
    # Read customer numbers from Excel file
    customer_numbers = read_customer_numbers_from_excel(excel_path)
    
    if not customer_numbers:
        logging.error("No customer numbers found in the Excel file or file could not be read.")
        input("Press Enter to exit...")
        return
    
    logging.info(f"Found {len(customer_numbers)} customer numbers in the Excel file.")
    
    # Contador de éxitos y fallos
    successful_customers = 0
    failed_customers = 0
    
    # Lista para guardar clientes que fallaron
    failed_customer_list = []
    
    automation = DellSalesPortalAutomation(webdriver_path, downloads_path)

    # Agregar timestamp al nombre del archivo de log consolidado para evitar sobrescribir
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    consolidated_log_path = f"C:\\Users\\L_Barnett\\OneDrive - Dell Technologies\\Documents\\My coding Projects\\DSA AUTOMATION\\Consolidated_Log_{timestamp}.txt"
    
    # Proceso de logging consolidado
    with open(consolidated_log_path, 'w') as log_file:
        # Procesar cada customer number
        for i, customer_id in enumerate(customer_numbers, 1):
            logging.info(f"===== Processing customer {i} of {len(customer_numbers)}: {customer_id} =====")
            log_file.write(f"===== Processing customer {i} of {len(customer_numbers)}: {customer_id} =====\n")
            
            try:
                # Inicializar nuevo navegador para cada cliente
                automation.initialize_driver()
                automation.navigate_to_sales_portal()
                
                # Procesar cliente
                automation.enter_customer_number(customer_id)
                automation.click_order_list_tab()
                automation.click_view_120_days()
                automation.adjust_date_range()
                automation.buttonsearch()
                automation.click_download_excel_button()
                
                # Procesar archivo descargado
                automation.process_downloaded_excel(customer_id)
                
                # Wait for download to complete
                time.sleep(5)
                
                # Incrementar contador de éxitos
                successful_customers += 1
                logging.info(f"Completed processing for customer: {customer_id}")
                log_file.write(f"Completed processing for customer: {customer_id}\n")
                
                # Cerrar el navegador después de la descarga exitosa
                automation.close_browser()
                
            except Exception as e:
                # Incrementar contador de fallos
                failed_customers += 1
                error_msg = f"Error processing customer {customer_id}: {str(e)}"
                logging.error(error_msg)
                log_file.write(error_msg + "\n")
                
                # Agregar a lista de clientes fallidos
                failed_customer_list.append(customer_id)
                
                # Cerrar el navegador en caso de error
                automation.close_browser()
    
    # Resumen final
    logging.info("----- PROCESSING SUMMARY -----")
    logging.info(f"Total Customers Processed: {len(customer_numbers)}")
    logging.info(f"Successful Customers: {successful_customers}")
    logging.info(f"Failed Customers: {failed_customers}")
    
    if failed_customer_list:
        logging.info("Failed Customer Numbers:")
        for customer in failed_customer_list:
            logging.info(customer)
    
    logging.info("Process complete.")
    
    # Crear un archivo de resumen
    summary_path = f"C:\\Users\\L_Barnett\\OneDrive - Dell Technologies\\Documents\\My coding Projects\\DSA AUTOMATION\\Processing_Summary_{timestamp}.txt"
    with open(summary_path, 'w') as summary_file:
        summary_file.write(f"Total Customers Processed: {len(customer_numbers)}\n")
        summary_file.write(f"Successful Customers: {successful_customers}\n")
        summary_file.write(f"Failed Customers: {failed_customers}\n")
        
        if failed_customer_list:
            summary_file.write("\nFailed Customer Numbers:\n")
            for customer in failed_customer_list:
                summary_file.write(f"{customer}\n")
    
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()