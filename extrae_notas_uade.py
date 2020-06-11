'''
Grades extractor from the UADE WebCampus2 website
'''

import sys
import os
from datetime import datetime

try:
    from bs4 import BeautifulSoup
    from selenium import webdriver
    from pandas import ExcelWriter, DataFrame

except ImportError as error:
    MODS = ['beautifulsoup4', 'pandas', 'selenium']
    print(f'Error: {error}')
    print('\n*Ocurrio un error al importar los modulos necesarios.')
    print('Asegurese de tener instalados por "pip install ..." los modulos:\n')
    for mod in MODS:
        print('-' + mod)
    sys.exit()

def opcion(texto='Si o no?: '):
    'Pregunta opcion, devuelve 1 para positivo, 0 para negativo'
    positivo = ['si', 's', '1']
    negativo = ['no', 'n', '0']
    while True:
        ingreso = input(texto).lower()
        if ingreso in positivo:
            resp = 1
            break
        if ingreso in negativo:
            resp = 0
            break
        print('Si o no?\n')
    return resp

def timer(func):
    'Prints function runtime'
    def wrapper():
        time1 = datetime.now()
        func()
        time2 = datetime.now()
        print(f'Runtime: {time2 - time1}')

    return wrapper

def wait_load(driver):
    'Waits for the new url to load before proceeding'
    driver.get(driver.current_url)

def login(driver, url, usr, psw):
    '''Logs in to the UADE webcampus profile

    Receives oper chromedriver.exe, login page url, user and password'''
    home_page = "HomeWC.aspx"
    if not psw.endswith('\n'):
        psw += '\n'
    driver.get(url)
    driver.find_element_by_id('ctl00_ContentPlaceHolderMain_txtUser').send_keys(usr)
    driver.find_element_by_id('ctl00_ContentPlaceHolderMain_txtClave1').send_keys(psw)
    wait_load(driver)
    if driver.current_url == url: # If no page change is detected
        raise PermissionError("'Login failed. Please verify that you're using the \
correct user/password.'")
    elif not driver.current_url.endswith(home_page):
        warnings_bypass(driver, home_page) # If logged in but not on homepage

def warnings_bypass(driver, homepage, tries=5):
    'Bypass the webcampus debt/ads screens'
    warning_buttons = [
        'ctl00_ContentPlaceHolderMain_SalteaPublicidad_Button1',
        'ctl00_ContentPlaceHolderMain_BtnContinuarWC']
    for i in range(tries):
        for button in warning_buttons:
            try:
                driver.find_element_by_id(button).click()
            except:
                pass
            finally:
                if driver.current_url.endswith(homepage):
                    wait_load(driver)
                    break
        else:
            continue
        break
    else:
        raise PermissionError("'Login failed. Could not make it past the \
warnings screen'")

def logout(driver):
    'Signs off UADE user'
    driver.find_element_by_id('ctl00_Top1_lnkCerrarSesion').click()

def kill_driver(driver):
    'Close the open chromedriver'
    driver.close()

def kill():
    'Kill the script exceution'
    print('\n\n* Se finalizara el programa.')
    sys.exit()

def extract_notes(table):
    '''Extract grades from the "cuatrimestre" page

    Table = HTML table grades section
    '''
    grades = []
    for grade in table.findAll('td', class_='td-texbox'):
        grades.append(grade.text)
    return grades

def extract_class_info(driver, url):
    '''Extract the info from the "cuatrimestre" page

    Returns classes (list), grades (matrix n_rows = len(classes))
    period (str with name of semester).
    '''
    driver.get(url)
    soup = BeautifulSoup(driver.page_source, features="html.parser")

    classes = []
    grades_list = []
    period = soup.find('tr', class_='td-ADMdoc-REG').text # Name of semester

    for classroom in soup.findAll('tr', class_="td-AULA-bkg"):  # Classroom table
        name = classroom.find('a').text.split('-')[0].strip()
        grades_table = classroom.find('td', class_='tabla-ID2').findAll('tr')[1]
        # Position 1 has the grades, position 0 the header
        grades = extract_notes(grades_table)

        classes.append(name)
        grades_list.append(grades)
    return classes, grades_list, period

def create_excel(class_matrix, header, student, file_name='Notas_UADE.xlsx'):
    '''Exports Class info to xlsx, requires pandas' ExcelWriter and DataFrame

    Receives:
    - class_matrix: each row containing: [class_info (names list), grades (matrix),
    sheet (str with desired excel sheet name), title (str with title for the first cel)
    - header: list containing the titles for the grades table
    - student: str, student name
    - file_name: optional, default = 'Notas_UADE.xlsx
    '''
    with ExcelWriter(file_name, mode='w+') as writer:
        for row in class_matrix:
            class_info, grades, sheet, title = row

            class_data = DataFrame(data=grades, index=class_info, columns=header)
            class_data.to_excel(writer, sheet_name=sheet, startrow=3)

            sheet = writer.sheets[sheet]
            sheet.cell(row=1, column=1).value = f'Alumno: {student}'
            sheet.cell(row=2, column=1).value = title
    writer.close()

def create_class_matrix(driver, links):
    ''' Create class info matrix: each row containing: [class_info (names list), grades (matrix),
    sheet (str with desired excel sheet name), title (str with title for the first cel)
    '''
    matrix = []
    for link in links:
        url = r'https://www.webcampus.uade.edu.ar/' + link
        clases, grades, period = extract_class_info(driver, url)
        page_name = period.split('-')[1].strip()

        if 'Cuatrimestre' in page_name:
            page_name.replace('Cuatrimestre', 'Cuatr.')

        matrix.append([clases, grades, page_name, period])
    return matrix

def same_type(matrix, col):
    "Checks that every item in matrix's column is the same type."
    prev_item = matrix[0][col]
    for i in range(1, len(matrix)):
        if not isinstance(matrix[i][col], type(prev_item)):
            break
        prev_item = matrix[i][col]
    else:
        return True
    return False

def uniform_matrix(matrix):
    "Determines if a given matrix contains uniform data types"
    cols = len(matrix[0])
    for i in range(cols):
        if not same_type(matrix, i):
            break
    else:
        return True
    return False

def name_extract(driver):
    """Receives chromedriver on webcampus' homepage
    Returns name, lastname. (both str)
    """
    soup = BeautifulSoup(driver.page_source, features="html.parser")
    name = soup.find('span', class_="TOPnombre").text.strip('Bienvenido')
    return name

def create_filename(name):
    'Creates the name for the excel file'
    lastname, name = [x.strip() for x in name.split(',')]
    f_name = name.split()[0].title()
    l_name = lastname.split()[0].title()
    file_name = f'{f_name}_{l_name}_Notas_UADE.xlsx'
    return file_name

def extract_links(driver):
    "Extract current and past semesters urls from webcampus' home"
    links = []
    soup = BeautifulSoup(driver.page_source, features="html.parser")

    for menu in soup.findAll('li', class_='rmItem'):
        if menu.text.split('\n')[0].lower() == 'mis cursos':
            for item in menu.findAll('li', class_="rsmItem"):
                item_name = item.text
                if 'cuatr' in item_name.lower() and item_name.endswith('Grado Monserrat'):
                    links.append(item.find('a')['href'])
    return links

@timer
def __main__():

    url = r'https://www.webcampus.uade.edu.ar/Login.aspx'
    driver_path = r"C:\Apps\chromedriver_win32\chromedriver.exe"
    while not os.path.isfile(driver_path):
        driver_website = r"https://sites.google.com/a/chromium.org/chromedriver/home"
        print(f'El webdriver de chrome no se encotro en la direccion "{driver_path}"\n')
        print(f'Link para la descarga del chrome web driver: "{driver_website}"')

        if opcion('\nDesea ingresar la direccion manualmente?: '):
            driver_path = input("Ingrese la direccion completa del chromedriver.exe: ").strip('"')
        else:
            kill()

    try:
        driver = webdriver.Chrome(driver_path)
        creds_file = 'cre.txt'
        while True:
            try:
                with open(creds_file, 'r') as creds:
                    usr, psw = creds.readline().split(';')

            except FileNotFoundError:
                print(f'No se encontro el archivo "{creds_file}".')
                print(f'\nSe recomienda crear un archivo llamado "{creds_file}" que contenga\
                    \nel usuario y la contraseña separados por un punto y coma ";"\n')

                if opcion("Desea ingresarlos manualmente?: "):
                    usr = input('Ingrese su usuario: ')
                    psw = input('Ingrese su contraseña: ')
                else:
                    kill()
            try:
                login(driver, url, usr, psw)
                break

            except PermissionError as error:
                print(f'\n* El inicio de sesion fallo: {error}')
                if opcion('Desea intentarlo de nuevo?: '):
                    continue
                kill_driver(driver)
                kill()

        st_name = name_extract(driver) # Studen name format lastname, name
        excel_filename = create_filename(st_name)
        links = extract_links(driver)
        if links != []:
            header = [
                'Parcial 1', 'Parcial 2', 'Recup', 'TP', 'Cursada',
                'Examen Final', 'Condicion Final', 'Asistencia'
                ]
            class_matrix = create_class_matrix(driver, links)
            logout(driver)
            kill_driver(driver)

            if uniform_matrix(class_matrix):
                create_excel(class_matrix, header, st_name, file_name=excel_filename)
            else:
                raise RuntimeError('Ocurrio un error al extraer la info de las cursadas.')
        else:
            raise RuntimeError('No se logro extraer ningun enlace a cuatrimestre cursado.')
    except PermissionError as error:
        print(f'> Ocurrio un error al momento de grabar el excel: {error}')
    except RuntimeError as error:
        print(f'> Ocurrio un error durante la ejecucion del programa: {error}')
    except Exception as error:
        print(f'> Error fatal inesperado: {error}')
    else:
        print('El programa fue ejecutado con exito!')

if __name__ == "__main__":
    __main__()
