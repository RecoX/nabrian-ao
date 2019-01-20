Attribute VB_Name = "Mod_General"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit

Public General_Connection_RenderRect As RECT
Public RenderRect As RECT

Public atacar As Integer
Public IsClan As Byte
Public NoRes As Boolean
Public NoGuia As Boolean
Public NoFps As Boolean
Public SombrasAC As Boolean
Public RetosAC As Boolean
Public CheckDobleAC As Boolean
Public DragAndDropAC As Boolean
Public ParticulasAC As Boolean
Public CreateDamageAC As Boolean
Public AurasAC As Boolean
Public Niebla As Boolean
Public HechizAc As Boolean
Public MinimapAc As Boolean
Public MeditacionesAZ As Boolean
Public activarnombresNpcs As Boolean
Public SkinGrafico As Integer
Public Desplazar As Boolean
Public vigilar As Boolean


Public rG(1 To 11, 1 To 3) As Byte

Public bO As Integer
Public bK As Long
Public bRK As Long
Public iplst As String
Public banners As String

Public bInvMod     As Boolean

Public bFogata As Boolean

Public bLluvia() As Byte

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal wIndx As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private lFrameLimiter As Long

Public lFrameModLimiter As Long
Public lFrameTimer As Long
Public sHKeys() As String

Public bFPS As Boolean
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32" (ByVal _
dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "Kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject _
As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" _
   (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Const PROCESS_TERMINATE = &H1
Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ACTIVE = &H103

Type Recompensa
    name As String
    Descripcion As String
End Type

Const GWL_STYLE = (-16)
Const Win_VISIBLE = &H10000000
Const Win_BORDER = &H800000
Const SC_CLOSE = &HF060&
Const WM_SYSCOMMAND = &H112

Dim ObjetoWMI As Object
Dim ProcesoACerrar As Object
Dim Procesos As Object
Public Recompensas(1 To 60, 1 To 3, 1 To 2) As Recompensa

Public Sub EstablecerRecompensas()

Recompensas(MINERO, 1, 1).name = "Fortaleza del Trabajador"
Recompensas(MINERO, 1, 1).Descripcion = "Aumenta la vida en 120 puntos."

Recompensas(MINERO, 1, 2).name = "Suerte de Novato"
Recompensas(MINERO, 1, 2).Descripcion = "Al morir hay 20% de probabilidad de no perder los minerales."

Recompensas(MINERO, 2, 1).name = "Destrucción Mágica"
Recompensas(MINERO, 2, 1).Descripcion = "Inmunidad al paralisis lanzado por otros usuarios."

Recompensas(MINERO, 2, 2).name = "Pica Fuerte"
Recompensas(MINERO, 2, 2).Descripcion = "Permite minar 20% más cantidad de hierro y la plata."

Recompensas(MINERO, 3, 1).name = "Gremio del Trabajador"
Recompensas(MINERO, 3, 1).Descripcion = "Permite minar 20% más cantidad de oro."

Recompensas(MINERO, 3, 2).name = "Pico de la Suerte"
Recompensas(MINERO, 3, 2).Descripcion = "Al morir hay 30% de probabilidad de que no perder los minerales (acumulativo con Suerte de Novato.)"


Recompensas(HERRERO, 1, 1).name = "Yunque Rojizo"
Recompensas(HERRERO, 1, 1).Descripcion = "25% de probabilidad de gastar la mitad de lingotes en la creación de objetos (Solo aplicable a armas y armaduras)."

Recompensas(HERRERO, 1, 2).name = "Maestro de la Forja"
Recompensas(HERRERO, 1, 2).Descripcion = "Reduce los costos de cascos y escudos a un 50%."

Recompensas(HERRERO, 2, 1).name = "Experto en Filos"
Recompensas(HERRERO, 2, 1).Descripcion = "Permite crear las mejores armas (Espada Neithan, Espada Neithan + 1, Espada de Plata + 1 y Daga Infernal)."

Recompensas(HERRERO, 2, 2).name = "Experto en Corazas"
Recompensas(HERRERO, 2, 2).Descripcion = "Permite crear las mejores armaduras (Armaduras de las Tinieblas, Armadura Legendaria y Armaduras del Dragón)."

Recompensas(HERRERO, 3, 1).name = "Fundir Metal"
Recompensas(HERRERO, 3, 1).Descripcion = "Reduce a un 50% la cantidad de lingotes utilizados en fabricación de Armas y Armaduras (acumulable con Yunque Rojizo)."

Recompensas(HERRERO, 3, 2).name = "Trabajo en Serie"
Recompensas(HERRERO, 3, 2).Descripcion = "10% de probabilidad de crear el doble de objetos de los asignados con la misma cantidad de lingotes."


Recompensas(TALADOR, 1, 1).name = "Músculos Fornidos"
Recompensas(TALADOR, 1, 1).Descripcion = "Permite talar 20% más cantidad de madera."

Recompensas(TALADOR, 1, 2).name = "Tiempos de Calma"
Recompensas(TALADOR, 1, 2).Descripcion = "Evita tener hambre y sed."


Recompensas(CARPINTERO, 1, 1).name = "Experto en Arcos"
Recompensas(CARPINTERO, 1, 1).Descripcion = "Permite la creación de los mejores arcos (Élfico y de las Tinieblas)."

Recompensas(CARPINTERO, 1, 2).name = "Experto de Varas"
Recompensas(CARPINTERO, 1, 2).Descripcion = "Permite la creación de las mejores varas (Engarzadas)."

Recompensas(CARPINTERO, 2, 1).name = "Fila de Leña"
Recompensas(CARPINTERO, 2, 1).Descripcion = "Aumenta la creación de flechas a 20 por vez."

Recompensas(CARPINTERO, 2, 2).name = "Espíritu de Navegante"
Recompensas(CARPINTERO, 2, 2).Descripcion = "Reduce en un 20% el coste de madera de las barcas."


Recompensas(PESCADOR, 1, 1).name = "Favor de los Dioses"
Recompensas(PESCADOR, 1, 1).Descripcion = "Pescar 20% más cantidad de pescados."

Recompensas(PESCADOR, 1, 2).name = "Pesca en Alta Mar"
Recompensas(PESCADOR, 1, 2).Descripcion = "Al pescar en barca hay 10% de probabilidad de obtener pescados más caros."


Recompensas(MAGO, 1, 1).name = "Pociones de Espíritu"
Recompensas(MAGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(MAGO, 1, 2).name = "Pociones de Vida"
Recompensas(MAGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(MAGO, 2, 1).name = "Vitalidad"
Recompensas(MAGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(MAGO, 2, 2).name = "Fortaleza Mental"
Recompensas(MAGO, 2, 2).Descripcion = "Libera el limite de mana máximo."

Recompensas(MAGO, 3, 1).name = "Furia del Relámpago"
Recompensas(MAGO, 3, 1).Descripcion = "Aumenta el daño base máximo de la Descarga Eléctrica en 10 puntos."

Recompensas(MAGO, 3, 2).name = "Destrucción"
Recompensas(MAGO, 3, 2).Descripcion = "Aumenta el daño base mínimo del Apocalipsis en 10 puntos."


Recompensas(NIGROMANTE, 1, 1).name = "Pociones de Espíritu"
Recompensas(NIGROMANTE, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(NIGROMANTE, 1, 2).name = "Pociones de Vida"
Recompensas(NIGROMANTE, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(NIGROMANTE, 2, 1).name = "Vida del Invocador"
Recompensas(NIGROMANTE, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

Recompensas(NIGROMANTE, 2, 2).name = "Alma del Invocador"
Recompensas(NIGROMANTE, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

Recompensas(NIGROMANTE, 3, 1).name = "Semillas de las Almas"
Recompensas(NIGROMANTE, 3, 1).Descripcion = "Aumenta el daño base mínimo de la magia en 10 puntos."

Recompensas(NIGROMANTE, 3, 2).name = "Bloqueo de las Almas"
Recompensas(NIGROMANTE, 3, 2).Descripcion = "Aumenta la evasión en un 5%."


Recompensas(PALADIN, 1, 1).name = "Pociones de Espíritu"
Recompensas(PALADIN, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(PALADIN, 1, 2).name = "Pociones de Vida"
Recompensas(PALADIN, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(PALADIN, 2, 1).name = "Aura de Vitalidad"
Recompensas(PALADIN, 2, 1).Descripcion = "Aumenta la vida en 5 puntos y el mana en 10 puntos."

Recompensas(PALADIN, 2, 2).name = "Aura de Espíritu"
Recompensas(PALADIN, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

Recompensas(PALADIN, 3, 1).name = "Gracia Divina"
Recompensas(PALADIN, 3, 1).Descripcion = "Reduce el coste de mana de Remover Paralisis a 250 puntos."

Recompensas(PALADIN, 3, 2).name = "Favor de los Enanos"
Recompensas(PALADIN, 3, 2).Descripcion = "Aumenta en 5% la posibilidad de golpear al enemigo con armas cuerpo a cuerpo."


Recompensas(CLERIGO, 1, 1).name = "Pociones de Espíritu"
Recompensas(CLERIGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(CLERIGO, 1, 2).name = "Pociones de Vida"
Recompensas(CLERIGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(CLERIGO, 2, 1).name = "Signo Vital"
Recompensas(CLERIGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(CLERIGO, 2, 2).name = "Espíritu de Sacerdote"
Recompensas(CLERIGO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(CLERIGO, 3, 1).name = "Sacerdote Experto"
Recompensas(CLERIGO, 3, 1).Descripcion = "Aumenta la cura base de Curar Heridas Graves en 20 puntos."

Recompensas(CLERIGO, 3, 2).name = "Alzamientos de Almas"
Recompensas(CLERIGO, 3, 2).Descripcion = "El hechizo de Resucitar cura a las personas con su mana, energía, hambre y sed llenas y cuesta 1.100 de mana."


Recompensas(BARDO, 1, 1).name = "Pociones de Espíritu"
Recompensas(BARDO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(BARDO, 1, 2).name = "Pociones de Vida"
Recompensas(BARDO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(BARDO, 2, 1).name = "Melodía Vital"
Recompensas(BARDO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(BARDO, 2, 2).name = "Melodía de la Meditación"
Recompensas(BARDO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(BARDO, 3, 1).name = "Concentración"
Recompensas(BARDO, 3, 1).Descripcion = "Aumenta la probabilidad de Apuñalar a un 20% (con 100 skill)."

Recompensas(BARDO, 3, 2).name = "Melodía Caótica"
Recompensas(BARDO, 3, 2).Descripcion = "Aumenta el daño base del Apocalipsis y la Descarga Electrica en 5 puntos."


Recompensas(DRUIDA, 1, 1).name = "Pociones de Espíritu"
Recompensas(DRUIDA, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(DRUIDA, 1, 2).name = "Pociones de Vida"
Recompensas(DRUIDA, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(DRUIDA, 2, 1).name = "Grifo de la Vida"
Recompensas(DRUIDA, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

Recompensas(DRUIDA, 2, 2).name = "Poder del Alma"
Recompensas(DRUIDA, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

Recompensas(DRUIDA, 3, 1).name = "Raíces de la Naturaleza"
Recompensas(DRUIDA, 3, 1).Descripcion = "Reduce el coste de mana de Inmovilizar a 250 puntos."

Recompensas(DRUIDA, 3, 2).name = "Fortaleza Natural"
Recompensas(DRUIDA, 3, 2).Descripcion = "Aumenta la vida de los elementales invocados en 75 puntos."


Recompensas(ASESINO, 1, 1).name = "Pociones de Espíritu"
Recompensas(ASESINO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(ASESINO, 1, 2).name = "Pociones de Vida"
Recompensas(ASESINO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(ASESINO, 2, 1).name = "Sombra de Vida"
Recompensas(ASESINO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(ASESINO, 2, 2).name = "Sombra Mágica"
Recompensas(ASESINO, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

Recompensas(ASESINO, 3, 1).name = "Daga Mortal"
Recompensas(ASESINO, 3, 1).Descripcion = "Aumenta el daño de Apuñalar a un 70% más que el golpe."

Recompensas(ASESINO, 3, 2).name = "Punteria mortal"
Recompensas(ASESINO, 3, 2).Descripcion = "Las chances de apuñalar suben a 25% (Con 100 skills)."


Recompensas(CAZADOR, 1, 1).name = "Pociones de Espíritu"
Recompensas(CAZADOR, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(CAZADOR, 1, 2).name = "Pociones de Vida"
Recompensas(CAZADOR, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(CAZADOR, 2, 1).name = "Fortaleza del Oso"
Recompensas(CAZADOR, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(CAZADOR, 2, 2).name = "Fortaleza del Leviatán"
Recompensas(CAZADOR, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(CAZADOR, 3, 1).name = "Precisión"
Recompensas(CAZADOR, 3, 1).Descripcion = "Aumenta la puntería con arco en un 10%."

Recompensas(CAZADOR, 3, 2).name = "Tiro Preciso"
Recompensas(CAZADOR, 3, 2).Descripcion = "Las flechas que golpeen la cabeza ignoran la defensa del casco."


Recompensas(ARQUERO, 1, 1).name = "Flechas Mortales"
Recompensas(ARQUERO, 1, 1).Descripcion = "1.500 flechas que caen al morir."

Recompensas(ARQUERO, 1, 2).name = "Pociones de Vida"
Recompensas(ARQUERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(ARQUERO, 2, 1).name = "Vitalidad Élfica"
Recompensas(ARQUERO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(ARQUERO, 2, 2).name = "Paso Élfico"
Recompensas(ARQUERO, 2, 2).Descripcion = "Aumenta la evasión en un 5%."

Recompensas(ARQUERO, 3, 1).name = "Ojo del Águila"
Recompensas(ARQUERO, 3, 1).Descripcion = "Aumenta la puntería con arco en un 5%."

Recompensas(ARQUERO, 3, 2).name = "Disparo Élfico"
Recompensas(ARQUERO, 3, 2).Descripcion = "Aumenta el daño base mínimo de las flechas en 5 puntos y el máximo en 3 puntos."


Recompensas(GUERRERO, 1, 1).name = "Pociones de Poder"
Recompensas(GUERRERO, 1, 1).Descripcion = "80 pociones verdes y 100 amarillas que no caen al morir."

Recompensas(GUERRERO, 1, 2).name = "Pociones de Vida"
Recompensas(GUERRERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(GUERRERO, 2, 1).name = "Vida del Mamut"
Recompensas(GUERRERO, 2, 1).Descripcion = "Aumenta la vida en 5 puntos."

Recompensas(GUERRERO, 2, 2).name = "Piel de Piedra"
Recompensas(GUERRERO, 2, 2).Descripcion = "Aumenta la defensa permanentemente en 2 puntos."

Recompensas(GUERRERO, 3, 1).name = "Cuerda Tensa"
Recompensas(GUERRERO, 3, 1).Descripcion = "Aumenta la puntería con arco en un 10%."

Recompensas(GUERRERO, 3, 2).name = "Resistencia Mágica"
Recompensas(GUERRERO, 3, 2).Descripcion = "Reduce la duración de la parálisis de un minuto a 45 segundos."


Recompensas(PIRATA, 1, 1).name = "Marejada Vital"
Recompensas(PIRATA, 1, 1).Descripcion = "Aumenta la vida en 20 puntos."

Recompensas(PIRATA, 1, 2).name = "Aventurero Arriesgado"
Recompensas(PIRATA, 1, 2).Descripcion = "Permite entrar a los dungeons independientemente del nivel."

Recompensas(PIRATA, 2, 1).name = "Riqueza"
Recompensas(PIRATA, 2, 1).Descripcion = "10% de probabilidad de no perder los objetos al morir."

Recompensas(PIRATA, 2, 2).name = "Escamas del Dragón"
Recompensas(PIRATA, 2, 2).Descripcion = "Aumenta la vida en 40 puntos."

Recompensas(PIRATA, 3, 1).name = "Magia Tabú"
Recompensas(PIRATA, 3, 1).Descripcion = "Inmunidad a la paralisis."

Recompensas(PIRATA, 3, 2).name = "Cuerda de Escape"
Recompensas(PIRATA, 3, 2).Descripcion = "Permite salir del juego en solo dos segundos."


Recompensas(LADRON, 1, 1).name = "Codicia"
Recompensas(LADRON, 1, 1).Descripcion = "Aumenta en 10% la cantidad de oro robado."

Recompensas(LADRON, 1, 2).name = "Manos Sigilosas"
Recompensas(LADRON, 1, 2).Descripcion = "Aumenta en 5% la probabilidad de robar exitosamente."

Recompensas(LADRON, 2, 1).name = "Pies sigilosos"
Recompensas(LADRON, 2, 1).Descripcion = "Permite moverse mientrás se está oculto."

Recompensas(LADRON, 2, 2).name = "Ladrón Experto"
Recompensas(LADRON, 2, 2).Descripcion = "Permite el robo de objetos (10% de probabilidad)."

Recompensas(LADRON, 3, 1).name = "Robo Lejano"
Recompensas(LADRON, 3, 1).Descripcion = "Permite robar a una distancia de hasta 4 tiles."

Recompensas(LADRON, 3, 2).name = "Fundido de Sombra"
Recompensas(LADRON, 3, 2).Descripcion = "Aumenta en 10% la probabilidad de robar objetos."

End Sub

Public Function DirGraficos() As String
DirGraficos = App.Path & "\" & "Graficos" & "\"
End Function

Public Function DirSound() As String
DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function
Public Function SD(ByVal N As Integer) As Integer

Dim auxint As Integer
Dim digit As Byte
Dim suma As Integer
auxint = N

Do
    digit = (auxint Mod 10)
    suma = suma + digit
    auxint = auxint \ 10

Loop While (auxint <> 0)

SD = suma

End Function

Public Function SDM(ByVal N As Integer) As Integer

Dim auxint As Integer
Dim digit As Integer
Dim suma As Integer
auxint = N

Do
    digit = (auxint Mod 10)
    
    digit = digit - 1
    
    suma = suma + digit
    
    auxint = auxint \ 10

Loop While (auxint <> 0)

SDM = suma

End Function

Public Function Complex(ByVal N As Integer) As Integer

If N Mod 2 <> 0 Then
    Complex = N * SD(N)
Else
    Complex = N * SDM(N)
End If

End Function

Public Function ValidarLoginMSG(ByVal N As Integer) As Integer
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(N)
AuxInteger2 = SDM(N)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function
Sub PlayWaveAPI(File As String)

On Error Resume Next
Dim rc As Integer

rc = sndPlaySound(File, SND_ASYNC)

End Sub


Sub Addtostatus(RichTextBox As RichTextBox, Text As String, Red As Byte, Green As Byte, Blue As Byte, Bold As Byte, Italic As Byte)

frmCargando.Status.SelStart = Len(RichTextBox.Text)
frmCargando.Status.SelLength = 0
frmCargando.Status.SelColor = RGB(Red, Green, Blue)

If Bold Then
    frmCargando.Status.SelBold = True
Else
    frmCargando.Status.SelBold = False
End If

If Italic Then
    frmCargando.Status.SelItalic = True
Else
    frmCargando.Status.SelItalic = False
End If

frmCargando.Status.SelText = Chr(13) & Chr(10) & Text

End Sub
Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
    With RichTextBox
        If Len(.Text) > 1500 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
    .SelBold = IIf(Bold, True, False)
    .SelItalic = IIf(Italic, True, False)
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        RichTextBox.Refresh
        
    End With
End Sub

Sub AddtoTextBox(textbox As textbox, Text As String)

textbox.SelStart = Len(textbox.Text)
textbox.SelLength = 0

textbox.SelText = Chr(13) & Chr(10) & Text

End Sub

Sub SaveGameini()

Config_Inicio.name = "BetaTester"
Config_Inicio.Password = "DammLamers"
Config_Inicio.Puerto = UserPort

Call EscribirGameIni(Config_Inicio)

End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function



Function CheckUserData(checkemail As Boolean) As Boolean

Dim loopc As Integer
Dim charAscii As Integer


If UserPassword = "" Then
    MsgBox "Ingrese la contraseña de su personaje.", vbInformation, "Password"
    Exit Function
End If

For loopc = 1 To Len(UserPassword)
    charAscii = Asc(mid$(UserPassword, loopc, 1))
    If LegalCharacter(charAscii) = False Then
        MsgBox "El password es inválido." & vbCrLf & vbCrLf & "Volvé a intentarlo otra vez." & vbCrLf & "Si el password es ese, verifica el estado del BloqMayús.", vbExclamation, "Password inválido"
        Exit Function
    End If
Next loopc

If UserName = "" Then
    MsgBox "Tenés que ingresar el Nombre de tu Personaje para poder Jugar.", vbExclamation, "Nombre inválido"
    Exit Function
End If

If Len(UserName) > 20 Then
    MsgBox ("El Nombre de tu Personaje debe tener menos de 20 letras.")
    Exit Function
End If

For loopc = 1 To Len(UserName)

    charAscii = Asc(mid$(UserName, loopc, 1))
    If LegalCharacter(charAscii) = False Then
        MsgBox "El Nombre del Personaje ingresado es inválido." & vbCrLf & vbCrLf & "Verifica que no halla errores en el tipeo del Nombre de tu Personaje.", vbExclamation, "Carácteres inválidos"
        Exit Function
    End If
    
Next loopc


CheckUserData = True

End Function
Sub UnloadAllForms()
On Error Resume Next
Dim mifrm As Form

For Each mifrm In Forms
    Unload mifrm
Next

End Sub

Function LegalCharacter(KeyAscii As Integer) As Boolean





If KeyAscii = 8 Then
    LegalCharacter = True
    Exit Function
End If


If KeyAscii < 32 Or KeyAscii = 44 Then
    LegalCharacter = False
    Exit Function
End If

If KeyAscii > 126 Then
    LegalCharacter = False
    Exit Function
End If


If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
    LegalCharacter = False
    Exit Function
End If


LegalCharacter = True

End Function

Sub SetConnected()
IntervaloConexionLogin = 0
frmPrincipal.DetectedCheats.Enabled = True
frmPrincipal.AntiExternos.Enabled = True

Connected = True

Call SaveGameini


Unload frmConectar


frmPrincipal.Label8.Caption = UserName

frmPrincipal.Visible = True

Call Audio.PlayWave(0, "login.wav")

'mciExecute "Close All"
'Call Audio.StopMidi

End Sub
Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

End Function
Public Function TiempoTranscurrido(ByVal Desde As Single) As Single

TiempoTranscurrido = Timer - Desde

If TiempoTranscurrido < -5 Then
    TiempoTranscurrido = TiempoTranscurrido + 86400
ElseIf TiempoTranscurrido < 0 Then
    TiempoTranscurrido = 0
End If

End Function
Sub MoveMe(Direction As Byte)

If CONGELADO Then Exit Sub

If Cartel Then Cartel = False

If ProxLegalPos(Direction) And Not UserMeditar And Not UserParalizado Then
    If TiempoTranscurrido(LastPaso) >= IntervaloPaso Then
        Call SendData("M" & Direction)
        Call DibujarMiniMapa
        LastPaso = Timer
        If Not UserDescansar Then
            Call EliminarChars(Direction)
            Call MoveCharByHead(UserCharIndex, Direction)
            Call MoveScreen(Direction)
            Call DoFogataFx
        End If
    End If
ElseIf CharList(UserCharIndex).Heading <> Direction Then Call SendData("CHEA" & Direction)
End If

frmPrincipal.mapa.Caption = NombreDelMapaActual & " [" & UserMap & " - " & UserPos.X & " - " & UserPos.Y & "]"

End Sub
Function ProxLegalPos(Direction As Byte) As Boolean

Select Case Direction
    Case NORTH
        ProxLegalPos = LegalPos(UserPos.X, UserPos.Y - 1)
    Case SOUTH
        ProxLegalPos = LegalPos(UserPos.X, UserPos.Y + 1)
    Case WEST
        ProxLegalPos = LegalPos(UserPos.X - 1, UserPos.Y)
    Case EAST
        ProxLegalPos = LegalPos(UserPos.X + 1, UserPos.Y)
End Select

End Function
Sub EliminarChars(Direction As Byte)
Dim X(2) As Integer
Dim Y(2) As Integer

Select Case Direction
    Case NORTH, SOUTH
        X(1) = UserPos.X - MinXBorder - 2
        X(2) = UserPos.X + MinXBorder + 2
    Case EAST, WEST
        Y(1) = UserPos.Y - MinYBorder - 2
        Y(2) = UserPos.Y + MinYBorder + 2
End Select

Select Case Direction
    Case NORTH
        Y(1) = UserPos.Y - MinYBorder - 3
        If Y(1) < 1 Then Y(1) = 1
        Y(2) = Y(1)
    Case EAST
        X(1) = UserPos.X + MinXBorder + 3
        If X(1) > 99 Then X(1) = 99
        X(2) = X(1)
    Case SOUTH
        Y(1) = UserPos.Y + MinYBorder + 3
        If Y(1) > 99 Then Y(1) = 99
        Y(2) = Y(1)
    Case WEST
        X(1) = UserPos.X - MinXBorder - 3
        If X(1) < 1 Then X(1) = 1
        X(2) = X(1)
End Select

For Y(0) = Y(1) To Y(2)
    For X(0) = X(1) To X(2)
        If X(0) > 6 And X(0) < 95 And Y(0) > 6 And Y(0) < 95 Then
            If MapData(X(0), Y(0)).CharIndex > 0 Then
                CharList(MapData(X(0), Y(0)).CharIndex).POS.X = 0
                CharList(MapData(X(0), Y(0)).CharIndex).POS.Y = 0
                MapData(X(0), Y(0)).CharIndex = 0
            End If
        End If
    Next
Next

End Sub
Public Sub ProcesaEntradaCmd(ByVal Datos As String)

If Len(Datos) = 0 Then Exit Sub


Select Case Left$(Datos, 1)
    Case "\", "/"
    
    Case Else
        Datos = ";" & Left$(frmPrincipal.modo, 1) & Datos

End Select

Call SendData(Datos)

End Sub
Public Sub ResetIgnorados()
Dim i As Integer

For i = 1 To UBound(Ignorados)
    Ignorados(i) = ""
Next

End Sub
Public Function EstaIgnorado(CharIndex As Integer) As Boolean
Dim i As Integer

For i = 1 To UBound(Ignorados)
    If Len(Ignorados(i)) > 0 And Ignorados(i) = CharList(CharIndex).Nombre Then
        EstaIgnorado = True
        Exit Function
    End If
Next

End Function
Sub CheckKeys()
On Error Resume Next
 
Static KeyTimer As Integer
 
If KeyTimer > 0 Then
    KeyTimer = KeyTimer - 1
    Exit Sub
End If
 
If Comerciando > 0 Then Exit Sub
       
If UserMoving = 0 Then
    If Not UserEstupido Then
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
            Call MoveMe(NORTH)
            Exit Sub
        End If
       
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
            Call MoveMe(EAST)
            Exit Sub
        End If
   
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
            Call MoveMe(SOUTH)
            Exit Sub
        End If
 
        If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
              Call MoveMe(WEST)
              Exit Sub
        End If
    Else
        Dim kp As Boolean
        kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
        If kp Then Call MoveMe(Int(RandomNumber(1, 4)))
    End If
End If
 
End Sub
Sub MoveScreen(Heading As Byte)
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer
Dim bx As Integer
Dim by As Integer

Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select

tX = UserPos.X + X
tY = UserPos.Y + Y



If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1

    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 8 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
Exit Sub
Stop
    
        
        Select Case FramesPerSecCounter
            Case Is >= 17
                lFrameModLimiter = 60
            Case 16
                lFrameModLimiter = 120
            Case 15
                lFrameModLimiter = 240
            Case 14
                lFrameModLimiter = 480
            Case 15
                lFrameModLimiter = 960
            Case 14
                lFrameModLimiter = 1920
            Case 13
                lFrameModLimiter = 3840
            Case 1
                lFrameModLimiter = 60 * 256
            Case 0
            
        End Select
    

    Call DoFogataFx
End If

End Sub
Function NextOpenChar()
Dim loopc As Integer

loopc = 1

Do While CharList(loopc).active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function
Public Function DirMapas() As String

DirMapas = App.Path & "\maps\"

End Function

Sub EliminarDatosMapa()
Dim X As Integer
Dim Y As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).CharIndex > 0 Then Call EraseChar(MapData(X, Y).CharIndex)
        MapData(X, Y).ObjGrh.GrhIndex = 0
        MapData(X, Y).ObjGrh.name = ""
    Next X
Next Y

End Sub
Public Function ReadFieldDarkFly2(ByVal iPos As Long, ByRef sText As String, ByVal charAscii As Long) As String
' Mismo que anterior con los parametros formales...
 
    '
    ' @ maTih.-
     
    Dim Read_Field()    As String
 
 
 
    'Creo un array temporal.
    Read_Field = Split(sText, ChrW$(charAscii))
' Mismo que antes con chrW
     
    If (iPos - 1) <= UBound(Read_Field()) Then
       'devuelve
       ReadFieldDarkFly2 = (Read_Field(iPos - 1))
    End If
     
    End Function
Public Function ReadFieldOptimizado(ByVal POS As Long, ByRef Text As String, ByVal SepASCII As Long) As String '13.0x
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************
 
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
 
    For i = 1 To POS
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, ChrW$(SepASCII), vbBinaryCompare)
    Next i
   
    If CurrentPos = 0 Then
        ReadFieldOptimizado = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadFieldOptimizado = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
   
End Function
Public Function NumeroApuesta(Numero As Integer) As String
Dim MiNum As Byte

Select Case Numero
    Case Is <= 36
        NumeroApuesta = "l " & Numero & "."
    Case 37
        NumeroApuesta = " los primeros 12."
    Case 38
        NumeroApuesta = " los segundos 12."
    Case 39
        NumeroApuesta = " los últimos 12."
    Case 40
        NumeroApuesta = " los primeros 18."
    Case 41
        NumeroApuesta = " los pares."
    Case 42
        NumeroApuesta = " los rojos."
    Case 43
        NumeroApuesta = " los negros."
    Case 44
        NumeroApuesta = " los impares."
    Case 45
        NumeroApuesta = " los últimos 18."
    Case Is <= 69
        MiNum = 3 * Fix((Numero - 46) / 2) + 2
        If Numero Mod 2 = 0 Then
            NumeroApuesta = "l semipleno " & MiNum - 1 & "-" & MiNum & "."
        Else
            NumeroApuesta = "l semipleno " & MiNum & "-" & MiNum + 1 & "."
        End If
    Case Is <= 102
        NumeroApuesta = "l semipleno " & Numero - 69 & "-" & Numero - 66 & "."
    Case Is <= 124
        MiNum = (3 * Fix((Numero - 101) / 2) - 1)
        If Numero Mod 2 = 1 Then MiNum = MiNum - 1
        NumeroApuesta = "l cuadro " & MiNum & "-" & MiNum + 1 & "-" & MiNum + 3 & "-" & MiNum + 4 & "."
    Case Is <= 136
        MiNum = 1 + 3 * (Numero - 125)
        NumeroApuesta = " la fila del " & MiNum & " al " & MiNum + 2 & "."
    Case Is <= 147
        MiNum = 1 + 3 * (Numero - 137)
        NumeroApuesta = " la calle del " & MiNum & " al " & MiNum + 5 & "."
    Case 148
        NumeroApuesta = " la primer columna."
    Case 149
        NumeroApuesta = " la segunda columna."
    Case 150
        NumeroApuesta = " la tercer columna."
End Select
        
End Function
Public Function PonerPuntos(Numero As Long) As String
Dim i As Integer
Dim Cifra As String

Cifra = Str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next

PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)

End Function
Function FileExist(File As String, FileType As VbFileAttribute) As Boolean

FileExist = Len(Dir$(File, FileType)) > 0

End Function

Sub WriteClientVer()

Dim hFile As Integer
    
hFile = FreeFile()
Open App.Path & "\init\Ver.bin" For Binary Access Write As #hFile
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)
Put #hFile, , CLng(777)

Put #hFile, , CInt(App.Major)
Put #hFile, , CInt(App.Minor)
Put #hFile, , CInt(App.Revision)

Close #hFile

End Sub

Sub ReNombrarAutoUpdate()

If FileExist(App.Path & "\NuevoUpdater.exe", vbNormal) Then
    If FileExist(App.Path & "\AutoUpdateClient.exe", vbNormal) Then Call Kill(App.Path & "\AutoUpdateClient.exe")
    Name App.Path & "\NuevoUpdater.exe" As App.Path & "\AutoUpdateClient.exe"
End If

End Sub
Public Function IsIp(ByVal ip As String) As Boolean

Dim i As Integer
For i = 1 To UBound(ServersLst)
    If ServersLst(i).ip = ip Then
        IsIp = True
        Exit Function
    End If
Next i

End Function

Public Sub InitServersList(ByVal Lst As String)

Dim NumServers As Integer
Dim i As Integer, Cont As Integer
i = 1

Do While (ReadFieldOptimizado(i, RawServersList, Asc(";")) <> "")
    i = i + 1
    Cont = Cont + 1
Loop

ReDim ServersLst(1 To Cont) As tServerInfo

For i = 1 To Cont
    Dim cur$
    cur$ = ReadFieldOptimizado(i, RawServersList, Asc(";"))
    ServersLst(i).ip = ReadFieldOptimizado(1, cur$, Asc(":"))
    ServersLst(i).Puerto = ReadFieldOptimizado(2, cur$, Asc(":"))
    ServersLst(i).Desc = ReadFieldOptimizado(4, cur$, Asc(":"))
    ServersLst(i).PassRecPort = ReadFieldOptimizado(3, cur$, Asc(":"))
Next i

CurServer = 1

End Sub
Sub CargarMensajesV()
Dim i As Integer
Dim File As String
Dim Formato As String
Dim NumMensajes As Integer

File = App.Path & "\Init\MensajesV.dat"

NumMensajes = Val(GetVar(File, "INIT", "NumMensajes"))

ReDim Mensajes(1 To NumMensajes) As Mensajito

For i = 1 To NumMensajes
    Mensajes(i).code = GetVar(File, "Mensaje" & i, "C")
    Mensajes(i).mensaje = GetVar(File, "Mensaje" & i, "M")
    Formato = GetVar(File, "Mensaje" & i, "F")
    Mensajes(i).Red = Val(ReadFieldOptimizado(1, Formato, Asc("-")))
    Mensajes(i).Green = Val(ReadFieldOptimizado(2, Formato, Asc("-")))
    Mensajes(i).Blue = Val(ReadFieldOptimizado(3, Formato, Asc("-")))
    Mensajes(i).Bold = Val(ReadFieldOptimizado(4, Formato, Asc("-")))
    Mensajes(i).Italic = Val(ReadFieldOptimizado(5, Formato, Asc("-")))
Next

Call SaveMensajes

End Sub
Function Transcripcion(Original As String) As String
Dim i As Integer, Char As Integer

For i = 1 To Len(Original)
    Char = Asc(mid$(Original, i, 1)) + 232 + i ^ 2
    Do Until Char < 255
        Char = Char - 255
    Loop
    Transcripcion = Transcripcion & Chr$(Char)
Next
    
End Function
Function Traduccion(Original As String) As String
Dim i As Integer, Char As Integer

For i = 1 To Len(Original)
    Char = Asc(mid$(Original, i, 1)) - 232 - i ^ 2
    Do Until Char > 0
        Char = Char + 255
    Loop
    Traduccion = Traduccion & Chr$(Char)
Next
    
End Function
Sub CargarMensajes()
Dim i As Integer, NumMensajes As Integer, Leng As Byte

Open App.Path & "\Init\Mensajes.dat" For Binary As #1
Seek #1, 1

Get #1, , NumMensajes

ReDim Mensajes(1 To NumMensajes) As Mensajito

For i = 1 To NumMensajes
    Mensajes(i).code = Space$(2)
    Get #1, , Mensajes(i).code
    Mensajes(i).code = Traduccion(Mensajes(i).code)
    
    Get #1, , Leng
    Mensajes(i).mensaje = Space$(Leng)
    Get #1, , Mensajes(i).mensaje
    Mensajes(i).mensaje = Traduccion(Mensajes(i).mensaje)
    
    Get #1, , Mensajes(i).Red
    Get #1, , Mensajes(i).Green
    Get #1, , Mensajes(i).Blue
    Get #1, , Mensajes(i).Bold
    Get #1, , Mensajes(i).Italic
Next

Close #1

End Sub
Sub SaveMensajes()
Dim i As Integer, File As String

File = App.Path & "\Init\Mensajes.dat"

Open File For Binary As #1
Seek #1, 1

Put #1, , CInt(UBound(Mensajes))
For i = 1 To UBound(Mensajes)
    Put #1, , Transcripcion(Mensajes(i).code)
    Put #1, , CByte(Len(Mensajes(i).mensaje))
    Put #1, , Transcripcion(Mensajes(i).mensaje)
    Put #1, , Mensajes(i).Red
    Put #1, , Mensajes(i).Green
    Put #1, , Mensajes(i).Blue
    Put #1, , Mensajes(i).Bold
    Put #1, , Mensajes(i).Italic
Next

Close #1

End Sub
Public Sub ActualizarInformacionComercio(Index As Integer)

Dim SR As RECT, DR As RECT
SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.bottom = 32

Select Case Index
    Case 0
        frmComerciar.Label1(0).Caption = "Sin precio"
        If OtherInventory(frmComerciar.List1(0).ListIndex + 1).Amount <> 0 Then
            frmComerciar.Label1(1).Caption = PonerPuntos(CLng(OtherInventory(frmComerciar.List1(0).ListIndex + 1).Amount))
        ElseIf OtherInventory(frmComerciar.List1(0).ListIndex + 1).name <> "Nada" Then
            frmComerciar.Label1(1).Caption = "Ilimitado"
        Else
            frmComerciar.Label1(1).Caption = 0
        End If
        
        frmComerciar.Label1(5).Caption = OtherInventory(frmComerciar.List1(0).ListIndex + 1).name
        frmComerciar.List1(0).ToolTipText = OtherInventory(frmComerciar.List1(0).ListIndex + 1).name
        
        Select Case OtherInventory(frmComerciar.List1(0).ListIndex + 1).ObjType
            Case 2
                frmComerciar.Label1(3).Caption = "Max Golpe:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxHit
                frmComerciar.Label1(4).Caption = "Min Golpe:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinHit
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Caption = "Arma:"
                frmComerciar.Label1(2).Visible = True
            Case 3
                frmComerciar.Label1(3).Visible = True
                      frmComerciar.Label1(3).Caption = "Defensa máxima: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef
                frmComerciar.Label1(4).Caption = "Defensa mínima: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinDef
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "Casco/Escudo/Armadura"
                If OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef = 0 Then
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Caption = "Esta ropa no tiene defensa."
                End If
                If OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef > 0 Then
                    frmComerciar.Label1(3).Visible = False
                    frmComerciar.Label1(4).Caption = "Defensa: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinDef & "/" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxDef
                End If
            Case 11
                frmComerciar.Label1(3).Caption = "Max Efecto:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxModificador
                frmComerciar.Label1(4).Caption = "Min Efecto:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinModificador
                
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "Min Efecto:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).TipoPocion
                Select Case OtherInventory(frmComerciar.List1(0).ListIndex + 1).TipoPocion
                    Case 1
                        frmComerciar.Label1(2).Caption = "Modifica Agilidad:"
                    Case 2
                        frmComerciar.Label1(2).Caption = "Modifica Fuerza:"
                    Case 3
                        frmComerciar.Label1(2).Caption = "Repone Vida:"
                    Case 4
                        frmComerciar.Label1(2).Caption = "Repone Mana:"
                    Case 5
                        frmComerciar.Label1(2).Caption = "- Cura Envenenamiento -"
                        frmComerciar.Label1(3).Visible = False
                        frmComerciar.Label1(4).Visible = False
                End Select
            Case 24
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Visible = False
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "- Hechizo -"
            Case 31
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "- Fragata -"
                frmComerciar.Label1(4).Caption = "Min/Max Golpe: " & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MinHit & "/" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).MaxHit
                frmComerciar.Label1(3).Caption = "Defensa:" & OtherInventory(frmComerciar.List1(0).ListIndex + 1).Def
                frmComerciar.Label1(4).Visible = True
            Case Else
                frmComerciar.Label1(2).Visible = False
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Visible = False
        End Select
        
        If OtherInventory(frmComerciar.List1(0).ListIndex + 1).PuedeUsar > 0 Then
            frmComerciar.Label1(6).Caption = "No podés usarlo ("
            Select Case OtherInventory(frmComerciar.List1(0).ListIndex + 1).PuedeUsar
                Case 1
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Genero)"
                Case 2
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Clase)"
                Case 3
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Facción)"
                Case 4
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Skill)"
                Case 5
                    frmComerciar.Label1(6).Caption = frmComerciar.Label1(6).Caption & "Raza)"
            End Select
        Else
            frmComerciar.Label1(6).Caption = ""
        End If
        
        If OtherInventory(frmComerciar.List1(0).ListIndex + 1).GrhIndex > 0 Then
            Call DrawGrhtoHdc(frmComerciar.Picture1.hDC, OtherInventory(frmComerciar.List1(0).ListIndex + 1).GrhIndex)
        Else
            frmComerciar.Picture1.Picture = LoadPicture()
        End If
        
    Case 1
        frmComerciar.Label1(0).Caption = "Sin precio"
        frmComerciar.Label1(1).Caption = PonerPuntos(UserInventory(frmComerciar.List1(1).ListIndex + 1).Amount)
        frmComerciar.Label1(5).Caption = UserInventory(frmComerciar.List1(1).ListIndex + 1).name

        frmComerciar.List1(1).ToolTipText = UserInventory(frmComerciar.List1(1).ListIndex + 1).name
        Select Case UserInventory(frmComerciar.List1(1).ListIndex + 1).ObjType
            Case 2
                frmComerciar.Label1(2).Caption = "Arma:"
                frmComerciar.Label1(3).Caption = "Max Golpe:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxHit
                frmComerciar.Label1(4).Caption = "Min Golpe:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinHit
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(4).Visible = True
            Case 3
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(3).Caption = "Defensa máxima: " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef
                frmComerciar.Label1(4).Caption = "Defensa mínima: " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinDef
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                frmComerciar.Label1(2).Caption = "Casco/Escudo/Armadura"
                If UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef = 0 Then
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Caption = "Esta ropa no tiene defensa."
                End If
                If UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef > 0 Then
                    frmComerciar.Label1(3).Visible = False
                    frmComerciar.Label1(4).Caption = "Defensa " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinDef & "/" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxDef
                End If
            Case 11
                frmComerciar.Label1(3).Caption = "Max Efecto:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxModificador
                frmComerciar.Label1(4).Caption = "Min Efecto:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinModificador
                
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Visible = True
                
                Select Case UserInventory(frmComerciar.List1(1).ListIndex + 1).TipoPocion
                    Case 1
                        frmComerciar.Label1(2).Caption = "Aumenta Agilidad"
                    Case 2
                        frmComerciar.Label1(2).Caption = "Aumenta Fuerza"
                    Case 3
                        frmComerciar.Label1(2).Caption = "Repone Vida"
                    Case 4
                        frmComerciar.Label1(2).Caption = "Repone Mana"
                    Case 5
                        frmComerciar.Label1(2).Caption = "- Cura Envenenamiento -"
                        frmComerciar.Label1(3).Visible = False
                        frmComerciar.Label1(4).Visible = False
                End Select
            Case 24
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Visible = False
                frmComerciar.Label1(2).Caption = "- Hechizo -"
                frmComerciar.Label1(2).Visible = True
            Case 31
                frmComerciar.Label1(3).Visible = True
                frmComerciar.Label1(4).Visible = True
                frmComerciar.Label1(2).Caption = "- Fragata -"
                frmComerciar.Label1(4).Caption = "Min/Max Golpe: " & UserInventory(frmComerciar.List1(1).ListIndex + 1).MinHit & "/" & UserInventory(frmComerciar.List1(1).ListIndex + 1).MaxHit
                frmComerciar.Label1(3).Caption = "Defensa:" & UserInventory(frmComerciar.List1(1).ListIndex + 1).Def
                frmComerciar.Label1(4).Visible = True
            frmComerciar.Label1(2).Visible = True
            Case Else
                frmComerciar.Label1(2).Visible = False
                frmComerciar.Label1(3).Visible = False
                frmComerciar.Label1(4).Visible = False
        End Select
        
        If UserInventory(frmComerciar.List1(1).ListIndex + 1).GrhIndex > 0 Then
            Call DrawGrhtoHdc(frmComerciar.Picture1.hDC, UserInventory(frmComerciar.List1(1).ListIndex + 1).GrhIndex)
        Else
            frmComerciar.Picture1.Picture = LoadPicture()
        End If
        
End Select

frmComerciar.Picture1.Refresh

End Sub
Sub TelepPorMapa(X As Long, Y As Long)
Dim Columna As Long, Fila As Long

Columna = Fix((X - 25) / 18)
Fila = Fix((Y - 18) / 18)

Call SendData("#$" & Columna & "," & Fila)

End Sub



Sub MainShell()
On Error Resume Next
'launcher
FormLauncher.Visible = False

Call WriteClientVer
 
NoRes = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana"))
NoFps = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "FPSLIBRE"))
NoGuia = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "GuiaJuego"))
SombrasAC = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Sombras"))
ParticulasAC = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Particulas"))
AurasAC = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Auras"))
HechizAc = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Hechiz"))
MeditacionesAZ = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Medit"))
Niebla = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Niebla"))
activarnombresNpcs = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "NombresNPC"))
Musica = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Music"))
CheckDobleAC = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "DobleClick"))
DragAndDropAC = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "DragAndDrop"))
SkinGrafico = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "SkinGrafico"))
MinimapAc = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "Minimap"))
CreateDamageAC = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "CreateDamage"))
RetosAC = Val(GetVar(App.Path & "/Init/Opciones.opc", "CONFIG", "RetosAC"))

If NoGuia = 0 Then
Shell "ErroresFIX.exe"
MsgBox ("Es la primera ves que ejecutas el juego se registraran librerías para solucionar errores dentro del juego.")
End If

FormLauncher.LabelVersion.Caption = VersionDelJuego

If NoFps = 0 Then
ActivadoFps = 1
FPslocos = 60
Else
ActivadoFps = 0
FPslocos = 15
End If
'launcher

If SeguridadActiva = True Then
If AoDefDebugger Then
Call AoDefAntiDebugger
End
End If

AoDefAntiShInitialize
AoDefOriginalClientName = "NabrianAO"
AoDefClientName = App.EXEName

If AoDefChangeName Then
Call AoDefClientOn
End
End If

If AoDefMultiClient Then
Call AoDefMultiClientOn
End
End If
End If

AddtoRichTextBox frmCargando.Status, "Cargando NabrianAO..", 255, 255, 255, 1, 1, False

 
ChDrive App.Path
ChDir App.Path
     

frmCargando.Show
frmCargando.Refresh
 

 
UserParalizado = False

AddtoRichTextBox frmCargando.Status, "Buscando servidor....", 200, 200, 200, 0, 0, True

AddtoRichTextBox frmCargando.Status, "OK!", 11, 213, 105, 1, 0, False

AddtoRichTextBox frmCargando.Status, "Buscando actualizaciones...", 200, 200, 200, 0, 0, True
Call frmCargando.Analizar
AddtoRichTextBox frmCargando.Status, "OK!", 11, 213, 105, 1, 0, False

'GM adm
rG(1, 1) = 227
rG(1, 2) = 235
rG(1, 3) = 8
'ciuda
rG(2, 1) = 0
rG(2, 2) = 128
rG(2, 3) = 255
'crimi
rG(3, 1) = 255
rG(3, 2) = 0
rG(3, 3) = 0
'nw
rG(4, 1) = 0
rG(4, 2) = 240
rG(4, 3) = 0
'NEUTRO
rG(5, 1) = 190
rG(5, 2) = 190
rG(5, 3) = 190
'Concilio de nabrian
rG(6, 1) = 255
rG(6, 2) = 128
rG(6, 3) = 190
'Consejo de banderbill
rG(7, 1) = 0
rG(7, 2) = 215
rG(7, 3) = 215
'Consejo negro
rG(8, 1) = 50
rG(8, 2) = 50
rG(8, 3) = 50
'SOPORTE
rG(9, 1) = 12
rG(9, 2) = 168
rG(9, 3) = 52
'EVENTOS
rG(10, 1) = 5
rG(10, 2) = 245
rG(10, 3) = 67
'SupeRADM
rG(11, 1) = 220
rG(11, 2) = 78
rG(11, 3) = 21
 
 
 
ReDim Ciudades(1 To NUMCIUDADES) As String
Ciudades(1) = "Ullathorpe"
Ciudades(2) = "Nix"
Ciudades(3) = "Banderbill"

ReDim CityDesc(1 To NUMCIUDADES) As String
CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
CityDesc(2) = "Nix es una gran ciudad. Edificada sobre la costa oeste del principal continente de Argentum."
CityDesc(3) = "Banderbill se encuentra al norte de Ullathorpe y Nix, es una de las ciudades más importantes de todo el imperio."

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"

ReDim ListaClases(1 To NUMCLASES) As String
ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Bandido"
ListaClases(9) = "Paladin"
ListaClases(10) = "Arquero"
ListaClases(11) = "Pescador"
ListaClases(12) = "Herrero"
ListaClases(13) = "Leñador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Pirata"

ReDim SkillsNames(1 To NUMSKILLS) As String
SkillsNames(1) = "Magia"
SkillsNames(2) = "Robar"
SkillsNames(3) = "Tacticas de combate"
SkillsNames(4) = "Combate con armas"
SkillsNames(5) = "Meditar"
SkillsNames(6) = "Apuñalar"
SkillsNames(7) = "Ocultarse"
SkillsNames(8) = "Supervivencia"
SkillsNames(9) = "Talar árboles"
SkillsNames(10) = "Defensa con escudos"
SkillsNames(11) = "Pesca"
SkillsNames(12) = "Mineria"
SkillsNames(13) = "Carpinteria"
SkillsNames(14) = "Herreria"
SkillsNames(15) = "Liderazgo"
SkillsNames(16) = "Domar animales"
SkillsNames(17) = "Armas de proyectiles"
SkillsNames(18) = "Wresterling"
SkillsNames(19) = "Navegacion"
SkillsNames(20) = "Sastrería"
SkillsNames(21) = "Comercio"
SkillsNames(22) = "Resistencia Mágica"

ReDim UserSkills(1 To NUMSKILLS) As Integer
ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"

 
AddtoRichTextBox frmCargando.Status, "Cargando Sonidos....", 200, 200, 200, 0, 0, True
lastTime = GetTickCount
    Call Audio.Initialize(frmPrincipal.hwnd, App.Path & "\wav\", App.Path & "\midi\")
    
AddtoRichTextBox frmCargando.Status, "Hecho", 11, 213, 105, 1, 0, False
 
 
   If MsgBox("¿Desea jugar en pantalla completa?", vbYesNo, "Cambio de resolución.") = vbYes Then
        NoRes = 0
          Call Resolution.SetResolution
        Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 0)
        Pantalla = True
    Else
        NoRes = 1
        Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 1)
        Pantalla = False
End If
 
 
ENDC = Chr(1)
 
UserMap = 1
 
Call CargarAnimsExtra
Call CargarArrayLluvia
Call CargarAnimArmas
Call CargarAnimEscudos
Call CargarMensajes
Call EstablecerRecompensas
 
Call InitTileEngine(frmPrincipal.renderer.hwnd, 32, 32, 13, 17)
 
AddtoRichTextBox frmCargando.Status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
 
Unload frmCargando
 
'Call Audio.PlayMIDI(App.Path & "\musicas\" & MIdi_Inicio & ".mid")
 

 
 
frmConectar.Visible = True
 
PrimeraVez = True
prgRun = True
Pausa = False
 
Dim ulttick As Long, esttick As Long
Dim timers(1 To 5) As Long
Dim loopc As Long
ScrollPixelsPerFrame = 7.25 'este valor sirve para modificar _
la caminata en la renderizacion y mas abajo
 
VelocidadCaminar = 0.029 'no hace falta explicarlo
 

Do While prgRun

    Call RefreshAllChars
    
    If EngineRun Then
        If frmPrincipal.WindowState <> 1 Then

        If UserMoving Then
            '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
                If ActivadoFps = 1 Then
                OffsetCounterX = (OffsetCounterX - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.X)))
                        If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
                Else
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrame * AddtoUserPos.X * timerTicksPerFrame
                        If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
                End If
        
            

            
            
            ElseIf AddtoUserPos.Y <> 0 Then
                If ActivadoFps = 1 Then
                OffsetCounterY = OffsetCounterY - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.Y))
                            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
                Else
               OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrame * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
                End If
    
            End If
        End If
    
    
    'If ActivadoFps = 0 Then
  '  Call ActualizarBarras
    'End If
    
    D3DDevice.BeginScene
     D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorXRGB(0, 0, 0), 1#, 0
            
            
            If UserCiego Then
                D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            Else
                RenderScreen UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY
            End If
            
            
            If ModoTrabajo Then Grh_Text_Render True, "MODO TRABAJO", 40, 1, D3DColorARGB(255, 255, 0, 0)
            If Cartel Then DibujarCartel
            Call Grh_Text_Render(False, FramesPerSec, 1, 1, D3DColorXRGB(200, 200, 200))
            If Dialogos.CantidadDialogos <> 0 Then Dialogos.MostrarTexto
            RenderSounds
            
         If Alphal > 1 Then DibujarNombreMapa 'Dibujado nombre mapas
         
        D3DDevice.Present RenderRect, ByVal 0, frmPrincipal.renderer.hwnd, ByVal 0
        D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.EndScene
    
    If ActivadoFps = 1 Then
            lFrameLimiter = GetTickCount
            FramesPerSecCounter = FramesPerSecCounter + 1
            timerElapsedTime = GetElapsedTime()
            timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
            particletimer = timerElapsedTime * 0.05
            Else
                 lFrameLimiter = GetTickCount
            FramesPerSecCounter = FramesPerSecCounter + 1
            timerElapsedTime = GetElapsedTime()
            timerTicksPerFrame = timerElapsedTime * VelocidadCaminar
            particletimer = timerElapsedTime * 0.05
            
            End If
        End If
    End If
    
If ActivadoFps = 1 Then
If (GetTickCount - lastTime > 20) Then
        If Not Pausa And frmPrincipal.Visible And Not frmForo.Visible Then
            CheckKeys
            lastTime = GetTickCount
        End If
    End If
Else
     If Not Pausa And frmPrincipal.Visible And Not frmForo.Visible Then
            CheckKeys
            lastTime = GetTickCount
        End If
End If
        
        
    If Not EngineRun Then
    renderconnect
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * 0.023
    particletimer = timerElapsedTime * 0.0005
    End If
    

    
    
    
        If ActivadoFps = 1 Then
While (GetTickCount - lFrameTimer) / FPslocos < FramesPerSecCounter
        Sleep 2
        Wend
        
                If GetTickCount - lFrameTimer > 1000 Then
            FramesPerSec = FramesPerSecCounter
            'If FPSFLAG Then frmPrincipal.Caption = "Nabrian AO" & " V " & App.Major & "." & App.Minor & "." & App.Revision & "-" & RandomNumber(2000, 3000)
            FramesPerSecCounter = 0
            lFrameTimer = GetTickCount
        End If
        
        Else
        
   While (GetTickCount - lFrameTimer) \ FPslocos < FramesPerSecCounter
      Sleep 5
    Wend
        
              If GetTickCount - lFrameTimer >= 1000 Then
        FramesPerSec = FramesPerSecCounter
       If FramesPerSec <> 0 Then ScrollPixelsPerFrame = 288 / 67
    
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
        End If
   
    
    ' ### I N T E R V A L O S ###
    esttick = GetTickCount
    For loopc = 1 To UBound(timers)
        timers(loopc) = timers(loopc) + (esttick - ulttick)
        
        If timers(1) >= tUs Then
            timers(1) = 0
            NoPuedeUsar = False
        End If
        
    Next loopc
    ulttick = GetTickCount
    
    DoEvents
Loop
 
EngineRun = False
frmCargando.Show
AddtoRichTextBox frmCargando.Status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
 
Call UnloadAllForms
'If NoRes = 0 Then Call Resolution.ResetResolution
Call DeInitTileEngine
End
 
'ManejadorErrores:
'    End
   
End Sub


Sub WriteVar(File As String, Main As String, Var As String, value As String)




writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(File As String, Main As String, Var As String) As String




Dim l As Integer
Dim Char As String
Dim sSpaces As String
Dim szReturn As String

szReturn = ""

sSpaces = Space(5000)


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

Public Function CheckMailString(ByRef sString As String) As Boolean
On Error GoTo errHnd:
Dim lPos As Long, lX As Long

lPos = InStr(sString, "@")
If (lPos <> 0) Then
    If Not InStr(lPos, sString, ".", vbBinaryCompare) > (lPos + 1) Then Exit Function

    For lX = 0 To Len(sString) - 1
        If Not lX = (lPos - 1) And Not CMSValidateChar_(Asc(mid$(sString, (lX + 1), 1))) Then Exit Function
    Next lX

    CheckMailString = True
End If
    
errHnd:

End Function
Private Function CMSValidateChar_(ByRef iAsc As Integer) As Boolean

CMSValidateChar_ = iAsc = 46 Or (iAsc >= 48 And iAsc <= 57) Or _
                    (iAsc >= 65 And iAsc <= 90) Or _
                    (iAsc >= 97 And iAsc <= 122) Or _
                    (iAsc = 95) Or (iAsc = 45)
                    
End Function
    Public Sub ShowSendTxt()
        If Not frmCantidad.Visible Then
            frmPrincipal.SendTxt.Visible = True
            Nopuede = 1
            frmPrincipal.SendTxt.SetFocus
        End If
    End Sub
Public Function AoDefCheatDetect(ByVal Chit As String)
Call SendData("BANEAME" & Chit)
MsgBox "Has sido echado por uso de " & Chit & " recuerda cerrar todo tipo de programa sospechoso para evitar que el juego te eche.", vbSystemModal, "Nabrian Security"
FrmAnticheat.Show , frmConectar
frmPrincipal.DetectedCheats.Enabled = False
frmPrincipal.AntiExternos.Enabled = False
End Function

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmPrincipal.renderer.ScaleWidth \ 64
    tY = UserPos.Y + viewPortY \ 32 - frmPrincipal.renderer.ScaleHeight \ 64
End Sub

Sub Base_Luz(Rojo As Byte, Verde As Byte, Azul As Byte)
'/////By Thusing/////
base_light = D3DColorXRGB(Rojo, Verde, Azul)
ColorLuz.R = Rojo
ColorLuz.G = Verde
ColorLuz.b = Azul
Light_Render_All
End Sub


Sub DibujarNombreMapa()
 
            If NombreMapaEspera > 0 Then
            If ActivadoFps Then
            NombreMapaEspera = NombreMapaEspera - (10)
            Else
            NombreMapaEspera = NombreMapaEspera - (3)
            End If
            End If
            
            If NombreMapaEspera <= 0 Then
            If ActivadoFps Then
            Alphal = Alphal - (10)
            Else
            Alphal = Alphal - (3)
            End If
            If Alphal < 50 Then Alphal = 0
            End If
            
           Call Grh_Text_Render(True, NombreDelMapaActual, 400, 1, D3DColorARGB(Val(Alphal), 255, 255, 255))

End Sub
 

Public Sub CargarParticulas()
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
 
    Dim StreamFile As String
    StreamFile = App.Path & "\init\" & "Particulas.ini"
 
    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
   
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As stream
   
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).X1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).Y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).X2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).Y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = 1 'Val(General_Var_Get(StreamFile, Val(LoopC), "AlphaBlend"))
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
       
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
       
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
       
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = Val(General_Field_Read(i, GrhListing, Asc(",")))
        Next i
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).R = Val(General_Field_Read(1, TempSet, Asc(",")))
            StreamData(loopc).colortint(ColorSet - 1).G = Val(General_Field_Read(2, TempSet, Asc(",")))
            StreamData(loopc).colortint(ColorSet - 1).b = Val(General_Field_Read(3, TempSet, Asc(",")))
        Next ColorSet
       
    Next loopc
End Sub
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0)
On Error Resume Next
 
Dim grh_list(1) As Long
grh_list(1) = 17220
 
'grh_list(1) = StreamData(ParticulaInd).grh_list(1)
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).R, _
            StreamData(ParticulaInd).colortint(0).G, _
            StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).R, _
            StreamData(ParticulaInd).colortint(1).G, _
            StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).R, _
            StreamData(ParticulaInd).colortint(2).G, _
            StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).R, _
            StreamData(ParticulaInd).colortint(3).G, _
            StreamData(ParticulaInd).colortint(3).b)
 
General_Particle_Create = Particle_Group_Create(X, Y, grh_list(), rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).X1, StreamData(ParticulaInd).Y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).X2, _
    StreamData(ParticulaInd).Y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)
 
End Function


Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: [URL='http://www.vbgore.com/GameClient.TileEngine.Engine_GetAngle']http://www.vbgore.com/GameClient.TileEn ... e_GetAngle[/URL]
'************************************************************
Dim SideA As Single
Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
Exit Function
 
End Function
 
 
 
Public Sub Crear_Flecha(ByVal AI As Integer, ByVal CI As Integer, ByVal Grh As Integer, ByVal R As Byte, fallo As Byte)
       
    Dim i As Integer
   
    For i = 1 To MaxFlecha
          If Flechas_list(i).EnUso = 0 Then Exit For
    Next
   
    If i = 0 Then i = 1
   
   
   
    Dim addfalloX As Byte
    Dim addFalloY As Integer
   
    If fallo <> 0 Then
   
    addfalloX = RandomNumber(-25, 25)
    addFalloY = RandomNumber(5, 25)
   
    Else
   
    addFalloY = -35
    'addfalloX = 16
   
    End If
   
    Flechas_list(i).xb = CharList(CI).POS.X * 32 + CharList(CI).MoveOffset.X + addfalloX
    Flechas_list(i).yb = CharList(CI).POS.Y * 32 + CharList(CI).MoveOffset.Y + addFalloY
    Flechas_list(i).Rotacion = R
    Flechas_list(i).X = CharList(AI).POS.X * 32 + CharList(AI).MoveOffset.X
    Flechas_list(i).Y = CharList(AI).POS.Y * 32 + CharList(AI).MoveOffset.Y '+ addFalloY
    Flechas_list(i).EnUso = 1
    InitGrh Flechas_list(i).Grh, Grh
   
    'Flechas_list(i).Angle = Engine_GetAngle(Flechas_list(i).xb, Flechas_list(i).Y, Flechas_list(i).xb, Flechas_list(i).yb)
   
End Sub
Public Sub DibujarMiniMapa()
   
Dim map_x As Long, map_y As Long
 
    For map_y = 1 To 100
        For map_x = 1 To 100
            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                SetPixel frmPrincipal.Minimap.hDC, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
            End If
        Next map_x
    Next map_y
   
    SetPixel frmPrincipal.Minimap.hDC, UserPos.X, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmPrincipal.Minimap.hDC, UserPos.X + 1, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmPrincipal.Minimap.hDC, UserPos.X - 1, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmPrincipal.Minimap.hDC, UserPos.X, UserPos.Y - 1, RGB(255, 0, 0)
    SetPixel frmPrincipal.Minimap.hDC, UserPos.X, UserPos.Y + 1, RGB(255, 0, 0)
 
    frmPrincipal.Minimap.Refresh
 
End Sub
Sub salirmsgbox()
        If MsgBox("¿Seguro que deseas salir?", vbYesNo, "Cambio de resolución") = vbYes Then
        Call SendData("/SALIR") 'z
        End
        Else
        End If
End Sub
Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long, outlinecolor As Long)
 
    Static box_rect As RECT
    Static Outline As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
   
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
   
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
   
    With box_rect
        .bottom = Y + Height - 1
        .Left = X + 1
        .Right = X + Width - 1
        .Top = Y + 1
    End With
   
    With Outline
        .bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With
   
   
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
   
   
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
 
   
End Sub

