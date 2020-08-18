VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'modulo que clase que se encarga de definir y conocer los objetos que se ven o no
'en el GUI de 3pm.
'además conoce su ubicación en coordenadas y tamaño

'lista de objetos que se pueden usar

'1 Lista de ritmos en uso
'2 Letras del abecedario
'3 Imagen de fondo. Contenedor global. Si uso esto podría tener un 3PM tipo Winamp
'necesito entonces tamaño diferenciado para el form y el picture por que puede ser
'que fijemos el tamaño a 800x600 y el tipo la use en 1024x768.
'4 Rectágulo donde entran los discos (a futuro que sea modificable por otro)
'5 VUMetros
'6 contenedor del video actual si es en monitor
'7 etiqueta con la cancion actual
'8 tiempo restante o actual de la canción
'9 lista de canciones pendientes
'10 cantidad de canciones pendientes
'11 imagen publicitaria en monitor
'12 credito disponible
'13 flechas de touch screen
'   a) derecha
'   b) izq
'   c) arriba
'   d) abajo
'   e) entrar a disco
'   f) comprar disco
'   g) comprar canción
'   h) escuchar canción
'   i) salir
'14 lista de canciones del disco actual (al entrar al disco o en todo momento)
'15 lista de discos en modo texto
'16 lista de canciones en modo texto del disco elegido (ver si puede ser la misma que la 14)
'17 publicidades de texto (inluye la obligatoria de tbr en lic simples) (son 2)
'18 imagen de inicio de 3pm
'19 barra de proceso de 3pm al iniciar el sistema
'20 identificador de la versión de 3pm


'la idea es que con cada objeto haya un contacto desde el programa
'habria que agrupar objetos parecidos para crear modelos
'permitir ademas que un mismo objeto pueda ser mostrado con mas de un modelo
'por ejemplo el indicador "CREDITO $4.80" puede ser un objeto simple (label)
'o con otro modelo que tenga imagen de fondo y texto. o con 2 label uno para la palabra
'credito y otro para el valor

'mientras mas objetos haya y mas representaciones tenga cada uno mayor será
'la flexibilidad de la interfase gráfica

'tipos de objetos

'1- Etiqueta simple
'   Propiedades: X,Y,W,H,Fuente,Transparente,BackColor,Forecolor,BorderStyle,alineacion
'2- Imagen: X,Y,W,H, transparente, borderstyle










