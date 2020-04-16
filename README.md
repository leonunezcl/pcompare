# Programa
Project Compare - Compara Proyectos Visual Basic

# Autor
Luis Leonardo Nuñez Ibarra. Año 2000 - 2003. email : leo.nunez@gmail.com. 

Chileno, casado , tengo 2 hijos. Aficionado a los videojuegos y el tenis de mesa. Mi primer computador fue un Talent MSX que me compro mi papa por alla por el año 1985. En el di mis primeros pasos jugando juegos como Galaga y PacMan y luego programando en MSX-BASIC. 

En la actualidad mi area de conocimiento esta referida a las tecnologias .NET con mas de 15 años de experiencia desarrollando varias paginas web usando asp.net con bases de datos sql server y Oracle. Integrador de tecnologias, desarrollo de servicios, aplicaciones de escritorio.

# Tipo de Proyecto
Project Compare es una aplicación para encontrar diferencias entre 2 proyectos Visual Basic. 

# Prologo
Regala un pescado a un hombre y le darás alimento para un día, enseñale a pescar y lo alimentarás para el resto de su vida (Proverbio Chino)

# Historia
Necesitaba una aplicación para comparar 2 versiones de un mismo proyecto. Si bien es cierto existen muchos programas que hacen una comparación "a nivel de archivo" no habia ninguna que lo hiciera tan al detalle como yo queria. Para esto y teniendo como base las rutinas de Project Explorer fue que desarrolle este utilitario. La idea final fue un comparador a nivel de variables, arreglos, apis de windows, archivos, procedimientos, funciones y todos aquellos elementos que me permitieran buscar diferencias entre un proyecto y otro.

# Archivos Necesarios
Este proyecto ocupa 5 componentes ActiveX 

- Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\SYSTEM\StdOle2.tlb#OLE Automation
- Reference=*\G{8B217740-717D-11CE-AB5B-D41203C10000}#1.0#0#C:\WINDOWS\SYSTEM\TLBINF32.DLL#TypeLib Information
- Reference=*\G{69EDFBA5-9FEC-11D5-89A4-F0FAEF3C8033}#1.0#0#C:\WINDOWS\SYSTEM\PVB_XMENU.DLL#PVB6 ActiveX DLL - Menu With Bitmaps !
- Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
- Object={3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0; RICHTX32.OCX

El archivo PVB_XMENU.DLL es un componente customizado para que los menus se puedan aplicar iconos y ayuda al momento de selección.

# Registro de los componentes ActiveX
Se debe realizar desde la linea de comando de windows regsvr32.exe [nombre del componente]
Para windows 10 necesitaras instalar con permisos de administrador. 

# Notas de los componentes ActiveX de Windows
Si obtienes error de licencia de componentes al momento de ejecutar el proyecto necesitaras instalar quizas la runtime de Visual Basic 5 (MSCVBM50.DLL) y bajar el archivo VB5CLI.EXE y VBUSC.EXE ambos disponibles en internet para descarga. Esto corregira los problemas de licencia de componentes de VB5.

# Desarrollo del proyecto
Como mencione anteriormente no tenia una aplicación para comparar 2 versiones de un proyecto tan al detalle como yo queria. Entonces teniendo la base del analizador de código project explorer y la idea de generar un comparador de proyectos fue que comenze desarrollando los algoritmos y rutinas para esta aplicación. 

De alguna manera nunca llegue a terminarla completamente pero deberia estar operativa casi al 100%

# Freeware
Por esos años mi intención fue ofrecerlo gratis a la comunidad Visual Basic que era bastante activa por esos años. Para esto levante un sitio web donde tenia varias otras aplicaciones que tambien habian sido creadas de la necesidad y que las distribuia de forma gratis.

# Palabras Finales
Espero que este proyecto que nacio de una necesidad personal sea usado con motivos de estudio y motivación. De como se pueden copiar las buenas ideas y mejorarlas. 
