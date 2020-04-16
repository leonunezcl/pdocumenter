# Programa
Project Documenter - Documenta proyectos Visual Basic

# Autor
Luis Leonardo Nuñez Ibarra. Año 2000 - 2003. email : leo.nunez@gmail.com. 

Chileno, casado , tengo 2 hijos. Aficionado a los videojuegos y el tenis de mesa. Mi primer computador fue un Talent MSX que me compro mi papa por alla por el año 1985. En el di mis primeros pasos jugando juegos como Galaga y PacMan y luego programando en MSX-BASIC. 

En la actualidad mi area de conocimiento esta referida a las tecnologias .NET con mas de 15 años de experiencia desarrollando varias paginas web usando asp.net con bases de datos sql server y Oracle. Integrador de tecnologias, desarrollo de servicios, aplicaciones de escritorio.

# Tipo de Proyecto
Project Documenter es una aplicación que se encarga de generar la documentación de un proyecto Visual Basic.

# Prologo
Regala un pescado a un hombre y le darás alimento para un día, enseñale a pescar y lo alimentarás para el resto de su vida (Proverbio Chino)

# Historia
Este utiitario en su origen no es de mi autoría. No recuerdo bien que buscaba en su momento y lo encontre en un sitio web. Estaba escrito en ingles en su totalidad y era perfecto para mis utilitarios personales. Traduci toda la interfaz al español y le hice algunas modificaciones a la interfaz. 

De este proyecto tambien saque varias ideas de como manejar la impresión de archivos y de como hacer impresión preliminar. Ademas tambien me sirvio para corregir algunos problemas que tenia con la lectura de proyectos de project explorer.

# Archivos Necesarios
Este proyecto ocupa 5 componentes ActiveX 

- Reference=*\G{69EDFBA5-9FEC-11D5-89A4-F0FAEF3C8033}#1.0#0#C:\WINDOWS\SYSTEM\PVB_XMENU.DLL#PVB6 ActiveX DLL - Menu With Bitmaps !
- Object={6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0; COMCTL32.OCX
- Object={BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0; MSOUTL32.OCX
- Object={BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0; TABCTL32.OCX
- Object={F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0; COMDLG32.OCX
- Object={3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0; RICHTX32.OCX
- Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX

El archivo PVB_XMENU.DLL es un componente customizado para que los menus se puedan aplicar iconos y ayuda al momento de selección.

# Registro de los componentes ActiveX
Se debe realizar desde la linea de comando de windows regsvr32.exe [nombre del componente]
Para windows 10 necesitaras instalar con permisos de administrador. 
