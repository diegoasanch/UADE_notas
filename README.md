# Exportador de notas del webcampus de UADE a xlsx #

Inicia sesion en tu cuenta de UADE webcampus y extrae la informacion de cuatrimestres cursados y notas obtenidas, y luego exportandola a
un archivo xlsx.

Requiere tener instalados los modulos:
    - BeautifulSoup4
    - Pandas
    - Selenium

Y la version de chromedriver.exe compatible con la version instalada de google chrome.
    - link para el webdriver: https://sites.google.com/a/chromium.org/chromedriver/home
  
Para la ejecucion del programa se sugiere tener en el mismo directorio un archivo llamado "cre.txt" que contenga las credenciales
de ingreso al webcampus en el formato: "usuario";"contraseña" aunque las mismas pueden ser ingresadas por medio de la terminal de 
python activa durante la ejecucion.
