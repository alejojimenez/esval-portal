import os
import time
import shutil
import rpa as r
from pathlib import Path
import shutil
import pandas as pd
#from domain.chrome_node import ChromeNode

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

from openpyxl import load_workbook
from datetime import datetime,timedelta

class Scraper_Esval():

    def __init__(self,url, email, password, driver_path):
        print(url, email, password, driver_path)
        self.url = url
        self.email = email
        self.password = password
        self.driver_path = driver_path

    def wait(self, seconds):
        return WebDriverWait(self.driver, seconds)

    def close(self):
        self.driver.close()
        self.driver = None

    def quit(self):
        self.driver.quit()
        self.driver = None
        
    def diccionario(self):
        dic_mes = {'01':'Enero','02':'Febrero',
            '03':'Marzo','04':'Abril',
            '05':'Mayo','06':'Junio',
            '07':'Julio','08':'Agosto',
            '09':'Septiembre','10':'Octubre',
            '11':'Noviembre','12':'Diciembre',
            }
        return dic_mes
        
    def get_downloads_folder(self):
        # Obtener la ruta del directorio de inicio del usuario actual
        home_dir = os.path.expanduser("~")

        # Combinar el directorio de inicio con la carpeta de descargas para obtener la ruta completa
        downloads_folder = os.path.join(home_dir, "Downloads")  # "Descargas"

        return downloads_folder
    
    def login(self):
    
        #driver_exe = 'C:\\roda\\esval-portal\\domain\\chromedriver.exe'
        driver_exe = './chromedriver.exe'

        print('Entrando en la funcion login...')
        print('----------------------------------------------------------------------')
        
        #Seteo variables
        email = self.email
        url = self.url
        driver_path = self.driver_path
        password  = self.password
        
        #Version 4.9.1 de selenium
        options = webdriver.ChromeOptions()    
        options.add_argument("--disable-notifications")
        self.driver = webdriver.Chrome(self.driver_path,options=options)

        self.driver.get(self.url)
        self.driver.maximize_window()
        self.driver.implicitly_wait(40)

        # Controlar evento alert() Notificacion
        try:
            alert = WebDriverWait(self.driver, 35).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.accept()
                    
        except:
            print("No se encontró ninguna alerta.")  

        #Botones ID y XPATH que utilizaremos para logear
        selector_username_input = 'rut'
        selector_password_input = 'contrasena'
        selector_ingreso_button = 'btn-primary'

        #Encontrar la linea de usuario y setear usuario
        intentos = 0
        set_usuario = True
        while (set_usuario):
                try:
                    print('Try en la funcion usuario...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    element_username = WebDriverWait(self.driver, 20).until(
                    EC.element_to_be_clickable((By.ID, selector_username_input)))
                    element_username.clear()
                    element_username.click()
                    element_username.send_keys(email)
                    set_usuario = False
                except:    
                    print('Exception en la funcion usuario...')
                    print('----------------------------------------------------------------------')
                    set_usuario = intentos <= 3  
            
        #Encontrar la linea de clave y setear clave
        intentos = 0
        set_clave = True
        while (set_clave):
                try:
                    print('Try en la funcion clave ...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    element_password= WebDriverWait(self.driver, 20).until(
                    EC.element_to_be_clickable((By.ID, selector_password_input)))
                    element_password.clear()
                    element_password.click()
                    element_password.send_keys(password)
                    time.sleep(2)
                    set_clave = False
                except:    
                    print('Exception en la funcion clave ...')
                    print('----------------------------------------------------------------------')
                    set_clave = intentos <= 3
        
        #Hacemos click en el boton de ingreso
        intentos = 0
        ingreso = True
        while (ingreso):
                try:
                    print('Try en el boton de ingreso ...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    boton_ingreso = WebDriverWait(self.driver, 20).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, selector_ingreso_button)))
                    boton_ingreso.click()
                    ingreso = False
                except:    
                    print('Exception en el boton de ingreso...')
                    print('----------------------------------------------------------------------')
                    ingreso = intentos <= 3        
        
        print('Ya logramos ingresar, ahora vamos a buscar las factruras')
        time.sleep(7)
        
    def scrapping_esval(self):
        print('Entrando en la funcion Scrapping...')
        print('----------------------------------------------------------------------')
        
        folder_path_config = './config/'
        
        # Especifica la ruta de tu archivo Excel
        excel_file = folder_path_config + "clientes.xlsx"

        # Especifica el nombre de la hoja en la que se encuentran los datos
        hoja_excel = "Hoja1"

        # Carga los datos de Excel en un DataFrame
        df = pd.read_excel(excel_file, sheet_name=hoja_excel)
        print('Dataframe ', df)
        print('----------------------------------------------------------------------')

        #Buscamos el menu de las boletas para hacer click
        intentos = 0
        menu_boletas= True
        while (menu_boletas):
            try:
                print('Try localizando el menu de las boletas ...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                boletas = WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.XPATH,'/html/body/app-root/app-contenedor-usuarios/div[3]/div/div/div[1]/app-lateral-menu/div/ul/li[2]/a')))
                boletas.click()
                menu_boletas = False
            except:    
                print('Exception en el menu de las boletas ...')
                print('----------------------------------------------------------------------')
                menu_boletas = intentos <= 3    
                    
        #Dentro del menu de boletas buscamos el boton de mis cobros
        intentos = 0
        mis_cobros = True
        while (mis_cobros):
            try:
                print('Try localizando mis cobros dentro del menu boletas ...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                cobros = WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.XPATH,'/html/body/app-root/app-contenedor-usuarios/div[3]/div/div/div[1]/app-lateral-menu/div/ul/li[2]/div/ul/li[1]/ul/li[1]/a')))
                cobros.click()
                mis_cobros = False
            except:    
                print('Exception localizando mis cobros dentro del menu boletas ...')
                print('----------------------------------------------------------------------')
                mis_cobros = intentos <= 3    

        #Hacemos doble click sobre expansion para obtener el largo completo de las boletas
        largo_descarga = 13
        
        i = 1            
        while i < largo_descarga+1:

            #Ahora que tenemos el largo completo volvemos a subir para identificar el primera fila y comenzar descarga
            if i == 6 or i==10:
                self.driver.execute_script("window.scrollBy(0, 200)")
                botone_expandir= self.driver.find_element(By.XPATH, '//*[@id="tablaFacturacion"]/div/button/span')
                time.sleep(10)
                botone_expandir.click()
                        
            #Obtenemos la fecha y dividimos dicho valor para tener mes y año            
            reintentos = 0            
            fecha_oficial = True
            while (fecha_oficial) and reintentos <= 5:
                try:
                    print('Try localizando fecha en la tabla ...', reintentos)
                    print('----------------------------------------------------------------------')
                    reintentos += 1
                    fecha_completa = self.driver.find_element(By.XPATH,f'//*[@id="tablaFacturacion"]/div/table/tbody/tr[{i}]/td[1]')
                    fecha_texto = fecha_completa.text
                    print('fecha_texto: ', fecha_texto)
                    partes = fecha_texto.split("/")
                    mes = str(partes[1])
                    año = str(partes[2])
                    print('Hemos localizando fecha en la tabla ...', reintentos)
                    print('----------------------------------------------------------------------')
                    fecha_oficial = False
                except:    
                    print('Exception localizando fecha en la tabla...')
                    print('----------------------------------------------------------------------')
                    reintentos += 1

            #Obtenemos el boton de descarga que haremos click           
            reintentos = 0            
            descarga_oficial = True
            while (descarga_oficial) and reintentos <= 5:
                try:
                    print('Try localizando boton descarga ...', reintentos)
                    print('----------------------------------------------------------------------')
                    reintentos += 1
                    descarga = self.driver.find_element(By.XPATH,f'//*[@id="tablaFacturacion"]/div/table/tbody/tr[{i}]/td[5]/button[3]')
                    descarga.click()
                    time.sleep(15)
                    descarga_oficial = False
                except:    
                    print('Exception en el boton descarga...')
                    print('----------------------------------------------------------------------')
                    reintentos += 1

            # Esperar hasta que el elemento esté presente en la página
            self.driver.implicitly_wait(45)
            descargado = True
            while descargado:                
                # Obtener la URL de la ventana emergente
                downloads_folder = self.get_downloads_folder()

                #Definimos ruta de destino del nuevo archivo
                folder_path = './input/'

                #Buscamos el archivo que se haya descargado que comience con el nombre boleta para poder moverlo a la carpeta input
                for filename in os.listdir(downloads_folder):
                    if filename.startswith("Boleta"):
                        nombre_archivo = filename
                print(nombre_archivo)
                
                #Split de nombre_archivopara extraer Nro. Boleta
                nombre, extension = nombre_archivo.split(".")

                # Utilizando isdigit() para extraer solo los números
                bill_number = "".join(filter(str.isdigit, nombre))                
                
                #Obtengo resultado del diccionario
                resultado_diccionario = self.diccionario()
                    
                # Cruce datos faltantes para ontener
                for index, row in df.iterrows():
                    
                    df_nro_cliente = df.loc[index, 'nro_cliente']
                    df_sucursal = df.loc[index, 'sucursal']
                    print('Nro. Cliente: ', df_nro_cliente, 'Sucursal: ', df_sucursal)
                    print('--------------------------------------------------------------------------')

                #Defino nuevo nombre para archivo que moveremos al Input
                nombre_archivo_nuevo = str(df_nro_cliente)+"_"+str(bill_number)+"_"+año+".pdf"
                    
                shutil.move(f'{downloads_folder}\\{nombre_archivo}',f'{folder_path}{nombre_archivo_nuevo}')
                descargado = False
              
                print("No se encontró el elemento con el id especificado...")
                print('----------------------------------------------------------------------')

                print('Conteo de documentos: ', i)
                print('----------------------------------------------------------------------')                

                print('----------------------------------------------------------------------')
                        
                time.sleep(20)
                print('pasamos al siguiente archivo')
            i +=1 