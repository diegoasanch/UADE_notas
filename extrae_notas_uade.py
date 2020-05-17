try:
    from bs4 import BeautifulSoup
    from selenium import webdriver
    from pandas import ExcelWriter, DataFrame
    from datetime import datetime
    import os

except ImportError:
    print('*Ocurrio un error al importar los modulos necesarios para este programa.')
    print('\nAsegurese de tener instalados por medio de "pip install ..." los siguientes modulos:\n\
-beautifulsoup4\n-selenium\n-pandas\n')
    exit()

def opcion(texto='Si o no?: '):
    'Pregunta opcion, devuelve 1 para positivo, 0 para negativo'
    si = ['si', 's', '1']
    no = ['no', 'n', '0']
    while True:
        op = input(texto).lower()
        if op in si:
            x = 1
            break
        elif op in no:
            x = 0
            break
    return x

def timer(func):
    'Prints function runtime'

    def wrapper():
        time1 = datetime.now()
        func()
        time2 = datetime.now()
        print(f'Runtime: {time2 - time1}')
        
    return wrapper

def login(driver, url, usr, psw):

    driver.get(url)
    driver.find_element_by_id('ctl00_ContentPlaceHolderMain_txtUser').send_keys(usr)
    driver.find_element_by_id('ctl00_ContentPlaceHolderMain_txtClave1').send_keys(psw)
    driver.get(driver.current_url)
    if driver.current_url == url:
        raise PermissionError('Login failed.')

def logout(driver):

    driver.find_element_by_id('ctl00_Top1_lnkCerrarSesion').click()

def kill():
    print('\n\n* Se finalizara el programa.')
    exit()

def notesExtract(table):
    grades = []
    for grade in table.findAll('td', class_='td-texbox'):
        grades.append(grade.text)
    return grades
    
def classInfoExtract(driver, url):

    driver.get(url)
    soup = BeautifulSoup(driver.page_source, features="html.parser")
    
    classes = []
    grades_list = []

    period = soup.find('tr', class_='td-ADMdoc-REG').text
    
    for classroom in soup.findAll('tr', class_="td-AULA-bkg"):
        name = classroom.find('a').text
        
        grades_table = classroom.find('td', class_='tabla-ID2').findAll('tr')[1]
        grades= notesExtract(grades_table)

        classes.append(name)
        grades_list.append(grades)
    return classes, grades_list, period

def createExcel(class_matrix, header, file_name='Notas_UADE.xlsx'):

    with ExcelWriter(file_name, mode='w+') as writer:
        for row in class_matrix:
            class_info, grades, sheet, title = row 

            class_data = DataFrame(data=grades, index=class_info, columns=header)
            class_data.to_excel(writer, sheet_name=sheet, startrow=3)

            sheet = writer.sheets[sheet]
            sheet.cell(row=1, column=1).value = title
    writer.close()

def classMatrixCreate(driver, links):
    matrix = []
    for link in links:
        url = r'https://www.webcampus.uade.edu.ar/' + link
        clases, grades, period = classInfoExtract(driver, url)
        page_name = period.split('-')[1].strip()

        if 'Cuatrimestre' in page_name: page_name.replace('Cuatrimestre', 'Cuatr.')
        
        matrix.append([clases, grades, page_name, period])
    return matrix

def sameType(matrix, col):
    prev_item = matrix[0][col]
    for i in range(1, len(matrix)):
        if type(matrix[i][col]) != type(prev_item):
            break
        prev_item = matrix[i][col]
    else:
        return True
    return False
    
def matrizxIsUniform(matrix):
    cols = len(matrix[0])
    for i in range(cols):
        if not sameType(matrix, i):
            break
    else:
        return True
    return False
        

def semestersExtract(driver):
    
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
        print(f'El webdriver de chrome no se encotro en la direccion predeterminada "{driver_path}"\n')
        print(r'Link para la descarga del chrome web driver: https://sites.google.com/a/chromium.org/chromedriver/home')

        if opcion('\nDesea ingresar la direccion manualmente?: '):
            driver_path = input("Ingrese la direccion completa del chromedriver.exe: ")
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
                print(f'\nSe recomienda crear una archivo llamado "{creds_file}" que contenga\
                    \nel usuario y la contraseña separados por un punto y coma ";"\n')
                
                if opcion("Desea ingresarlos manualmente?: "):
                    usr = input('Ingrese su usuario: ')
                    psw = input('Ingrese su contraseña: ')
                else:
                    kill()
            finally:
                if not psw.endswith('\n'): psw += '\n'
            try:
                login(driver, url, usr, psw)
                break
            except PermissionError as e:
                print(f'\n* El inicio de sesion fallo: {"str(e)"}')
                if opcion('Desea intentarlo de nuevo?: '):
                    continue
                else:
                    kill()

        links = semestersExtract(driver)
        if links != []:
            header =  ['Parcial 1', 'Parcial 2', 'Recup', 'TP', 'Cursada', 'Examen Final', 'Condicion Final', 'Asistencia']
            class_matrix = classMatrixCreate(driver, links)
            logout(driver)

            if matrizxIsUniform(class_matrix):
                createExcel(class_matrix, header)
            else:
                raise RuntimeError('Ocurrio un error al extraer la info de las cursadas.')
        else:
            raise RuntimeError('No se logro extraer ningun enlace a cuatrimestre cursado.')

    except RuntimeError as e:
        print(f'> Ocurrio un error durante la ejecucion del programa:  "{str(e)}"')
    except Exception as e:
        print(f'> Error fatal inesperado: "{str(e)}"')
    else:
        print('El programa fue ejecutado con exito!')

if __name__ == "__main__":
    __main__()
