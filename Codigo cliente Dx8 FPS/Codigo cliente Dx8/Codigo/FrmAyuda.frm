VERSION 5.00
Begin VB.Form FrmAyuda 
   BorderStyle     =   0  'None
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      MouseIcon       =   "FrmAyuda.frx":0000
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
   End
End
Attribute VB_Name = "FrmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'F�nixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Private Sub Form_GotFocus()

Select Case Ayuda
    Case 1
        Select Case SubAyuda
            Case 1
                FrmAyuda.Label1.Caption = "Ser fiel a la Alianza" & vbCrLf & vbCrLf & "Al elegir ser fiel a la Alianza, tendr�s una serie de beneficios pero al mismo tiempo perder�s algunas posibilidades." & vbCrLf & vbCrLf & ". No podr�s atacar ni ser atacado por otros que tambi�n le hayan jurado fidelidad a la Alianza, as� como tampoco a sus mascotas (salvo que sea un clan con el cual se encuentren en guerra)." & _
                vbCrLf & ". Podr�s atacar tanto a aquellos que no le hayan jurado fidelidad a la Alianza, sin importar si le juraron fidelidad a Lord Thek o si decidieron no ser fieles a ninguno de los dos. De todas formas, la Alianza tendr� en cuenta para sus recompensas s�lo las muertes de quienes son fieles a Lord Thek." & vbCrLf & ". No podr�s atacar a los guardias de la Alianza que vigilan d�a y noche las ciudades de la Alianza del Nabrian." & vbCrLf & ". Si en alg�n momento decides dejar de serle fiel a la Alianza, los guardias de las ciudades te perseguir�n. Adem�s, la Alianza ya no te permitir� volver con �l nunca." & vbCrLf & ". Una vez que elegiste jurarle fidelidad a la Alianza, ya no podr�s jurarle fidelidad al Lord Thek jam�s, por m�s que hayas abandonado tu anterior vida." & vbCrLf & ". Tendr�s la posibilidad de pertenecer a la Alianza del Nabrian."
            
            Case 2
                FrmAyuda.Label1.Caption = "Ser fiel a Lord Thek" & vbCrLf & vbCrLf & "Al elegir ser fiel a Lord Thek obtendr�n varios beneficios aunque tambi�n perder�s algunas posibilidades que otorgan las otras elecciones." & vbCrLf & vbCrLf & ". No podr�s atacar ni ser atacado por quienes tambi�n le hayan jurado fidelidad a Lord Thek, as� como tampoco a sus mascotas (salvo que sea un clan con el cual se encuentren en guerra)." & _
                vbCrLf & ". Podr�s atacar tanto a aquellos que no le hayan jurado fidelidad a Lord Thek, sin importar si le juraron fidelidad a la Alianza o si decidieron no ser fieles a ninguno de los dos. De todas formas, la Alianza tendr� en cuenta para sus recompensas s�lo las muertes de quienes son fieles a Lord Thek." & vbCrLf & ". No podr�s atacar a los guardias de Lord Thek que vigilan d�a y noche las ciudades de la Legi�n Oscura." & vbCrLf & ". Si en alg�n momento decides dejar de serle fiel a Lord Thek, los guardias de las ciudades te perseguir�n. Adem�s Lord Thek ya no te permitir� volver con �l nunca." & vbCrLf & ". Una vez que elegiste jurarle fidelidad a Lord Thek, ya no podr�s jurarle fidelidad a la Alianza jam�s, por m�s que hayas abandonado tu anterior vida." & vbCrLf & ". Tendr�s la posibilidad de pertenecer a la Legi�n Oscura."
            
            Case 3
                FrmAyuda.Label1.Caption = "Mantenerse neutral" & vbCrLf & vbCrLf & "Al elegir ser neutral, est�s rechazando la oferta tanto de la Alianza como de Lord Thek de jurarles fidelidad." & vbCrLf & vbCrLf & ". Formas entonces tus propias reglas, puedes atacar y ser atacado por cualquier habitantes de estas tierras." & _
                vbCrLf & ". Podr�s m�s adelante jurarle fidelidad a la Alianza o a Lord Thek, pero para eso debes haber matado a menos personas que le hayan jurado fidelidad a quien deseas unirte que las que lo hicieron con el enemigo." & vbCrLf & ". Si atacas a los guardias de cualquier facci�n, los mismos te responder�n." & vbCrLf & ". Si en alg�n momento decides jurar fidelidad y luego te arrepientes y abandonas, los guardias de esa facci�n te perseguir�n." & vbCrLf & ". Una vez que elegiste jurarle fidelidad a la Alianza o a Lord Thek, si decides hacerlo, aunque lo abandones ya no podr�s jurar fidelidad al ej�rcito enemigo." & vbCrLf & vbCrLf & ". No podr�s formar parte de ning�n ej�rcito (Alianza del Nabrian / Legi�n Oscura) mientras seas neutral, as� como tampoco se otorgar�n recompensas que s� reciben ellos."
        End Select
    
    Case 0
        Select Case MiClase
             Case CIUDADANO
                 FrmAyuda.Label1.Caption = "Trabajador" & vbCrLf & vbCrLf & "Los trabajadores son todos aquellos que se dedican a fabricar items y conseguir alguna retribuci�n por ellos: mineros, herreros, sastres, pescadores, taladores y carpinteros." & vbCrLf & vbCrLf & "Luchador" & vbCrLf & vbCrLf & "Los luchadores son el resto de las clases, usen o no mana. Se incluyen aqu� adem�s a los ladrones y a los navegantes."
            
             Case TRABAJADOR
                 FrmAyuda.Label1.Caption = "Experto en minerales" & vbCrLf & vbCrLf & "Eligiendo esta profesi�n, posteriormente podr�s dedicarte a la extracci�n de minerales (miner�a) y su eventual trabajo de los mismos (herrer�a)." & vbCrLf & vbCrLf & "Experto en uso de madera" & vbCrLf & vbCrLf & "Siguiendo el oficio del uso de madera podremos posteriormente dedicarnos a la tala de �rboles (Talador) o el trabajo de la madera obtenida (Carpintero)." & vbCrLf & vbCrLf & "Experto en pesca" & vbCrLf & vbCrLf & "Personajes que por medio de una ca�a o una red de pesca pueden obtener oro vendiendo grandes cantidades de pescados." & vbCrLf & vbCrLf & "Sastre" & vbCrLf & vbCrLf & "Clase trabajadora que se encarga de hacer diferentes ropajes."
            
             Case EXPERTO_MINERALES
                 FrmAyuda.Label1.Caption = "Minero" & vbCrLf & vbCrLf & "Siendo minero podremos extraer diferentes minerales (hierro, plata, oro) y transformarlos en lingotes para su posterior utilizaci�n." & vbCrLf & vbCrLf & "Herrero" & vbCrLf & vbCrLf & "Dedic�ndonos al negocio de la herrer�a podremos trabajar lingotes para la fabricaci�n de armas y/o armaduras."
            
             Case EXPERTO_MADERA
                 FrmAyuda.Label1.Caption = "Le�ador" & vbCrLf & vbCrLf & "Siendo le�adores podremos talar �rboles y de esta forma conseguir le�os de madera que podr�n ser vendidos o trabajados." & vbCrLf & vbCrLf & "Carpintero" & vbCrLf & vbCrLf & "El carpintero es la clase que trabaja los le�os para poder fabricar diversos objetos como flechas, arcos o barcas."
         
             Case LUCHADOR
                 FrmAyuda.Label1.Caption = "Con uso de mana" & vbCrLf & vbCrLf & "Los personajes con uso de mana son aquellos luchadores que conocen las artes m�gicas: mago, nigromante, palad�n, cl�rigo, bardo, druida, cazador, asesino." & vbCrLf & vbCrLf & "Sin uso de mana" & vbCrLf & vbCrLf & "Quienes no usan mana son aquellos que no utilizan las artes m�gicas: guerrero, arquero, ladr�n y los piratas. Los piratas sirven especialmente para guardar tesoros, y no para la lucha."
             
             Case CON_MANA
                 FrmAyuda.Label1.Caption = "Hechicero" & vbCrLf & vbCrLf & "Utilizan exclusivamente hechizos para lanzar ataques. Siendo hechizero podr�s posteriormente transformarte en mago o nigromante." & vbCrLf & vbCrLf & "Orden Sagrada" & vbCrLf & vbCrLf & "Para integrar la Orden Sagrada no es necesario ser Ciudadano." & vbCrLf & "Eligiendo este camino podremos posteriormente transformarnos en paladines o cl�rigos." & vbCrLf & vbCrLf & "Naturalista" & vbCrLf & vbCrLf & "Transform�ndonos en Naturalistas podremos posteriormente convertirnos en bardos o druidas." & vbCrLf & vbCrLf & "Sigiloso" & vbCrLf & vbCrLf & "Eligiendo esta ramificaci�n podremos transformarnos posteriormente en asesinos o cazadores."
            
             Case HECHICERO
                 FrmAyuda.Label1.Caption = "Mago" & vbCrLf & vbCrLf & "El mago es un personaje de poca vida que utiliza la Furia del F�nix (mejor hechizo de mago) y puede llegar a tener m�s de 2000 de mana." & vbCrLf & vbCrLf & "Nigromante" & vbCrLf & vbCrLf & "El nigromante es un personaje de poca vida que puede usar el mejor hechizo de invocacion y llegar a tener como m�ximo 2000 de mana."
             
             Case ORDEN
                 FrmAyuda.Label1.Caption = "Palad�n" & vbCrLf & vbCrLf & "Los paladines son personajes que pueden lanzar hechizos de poca mana y a su vez utilizar diversas armaduras y armas para eliminar a sus enemigos." & vbCrLf & vbCrLf & "Cl�rigo" & vbCrLf & vbCrLf & "Los cl�rigos son personajes con una cantidad de mana superior al palad�n y un golpe con armas inferior."
             
             Case NATURALISTA
                 FrmAyuda.Label1.Caption = "Bardo" & vbCrLf & vbCrLf & "Los bardos son personajes de mana media y gran evasi�n de golpes, pero de golpe bajo." & vbCrLf & vbCrLf & "Druida" & vbCrLf & vbCrLf & "El druida es un personaje de mana media y golpe con arco medio que mantiene estrechos vinculos con la naturaleza."
             
             Case SIGILOSO
                 FrmAyuda.Label1.Caption = "Asesino" & vbCrLf & vbCrLf & "Personajes expertos en apu�alar (especialmente de ser elfo oscuro), de vida media y poca mana." & vbCrLf & vbCrLf & "Cazador" & vbCrLf & vbCrLf & "Los cazadores son personajes que pueden utilizar hechizos que requieran poca mana y a su vez se especializan en el uso de arco."
                  
             Case SIN_MANA
                 FrmAyuda.Label1.Caption = "Bandido" & vbCrLf & vbCrLf & "Son los que rompen las leyes y atacan a los jugadores desprevenidos (luego elegiras entre ser Ladr�n o Pirata)." & vbCrLf & vbCrLf & "Caballero" & vbCrLf & vbCrLf & "Personajes fuertes y valientes, los que tienen mayor potencia de ataque tanto en la lucha cuerpo a cuerpo como en el uso del arco y flecha (puedes transformate luego en Arquero o Guerrero)."
                  
             Case BANDIDO
                 FrmAyuda.Label1.Caption = "Pirata" & vbCrLf & vbCrLf & "El pirata es la clase navegadora por excelencia." & vbCrLf & vbCrLf & "Ladr�n" & vbCrLf & vbCrLf & "Roba una cantidad de oro igual a 50 veces su nivel con mayores probabilidades y pueden robar en ciudades."
        
             Case CABALLERO
                 FrmAyuda.Label1.Caption = "Guerrero" & vbCrLf & vbCrLf & "Personajes de gran vitalidad y fuerza que puede llegar a pegar grandes golpes o apu�aladas muy poderosas para eliminar a sus enemigos." & vbCrLf & vbCrLf & "Arquero" & vbCrLf & vbCrLf & "Son los que mejor dominan el uso del arco y flecha, teniendo una gran cantidad de vida."
      
    End Select
End Select

End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "masinfo.gif")

End Sub
Private Sub Form_LostFocus()

FrmAyuda.Hide

End Sub
Private Sub Image1_Click()

Unload Me

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving = False And Button = vbLeftButton Then
      Dx3 = X
      dy = Y
      bmoving = True
   End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving And ((X <> Dx3) Or (Y <> dy)) Then
      Move Left + (X - Dx3), Top + (Y - dy)
   End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      bmoving = False
   End If
End Sub

