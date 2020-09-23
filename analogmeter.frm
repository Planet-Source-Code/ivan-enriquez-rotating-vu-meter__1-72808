VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Mobile Meter Displays"
   ClientHeight    =   8760
   ClientLeft      =   1110
   ClientTop       =   1695
   ClientWidth     =   11445
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "analogmeter.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   763
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FF00&
      Height          =   1335
      Left            =   3480
      ScaleHeight     =   63.75
      ScaleMode       =   2  'Point
      ScaleWidth      =   231.75
      TabIndex        =   2
      Top             =   6960
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   4215
      Left            =   1680
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   1
      Top             =   1560
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   6240
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   0
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label8 
      Caption         =   "Start from this side |------->"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   8400
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "Meter Centered"
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Center off limits"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "CounterClockWise"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Clockwise"
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   584
      X2              =   584
      Y1              =   56
      Y2              =   88
   End
   Begin VB.Label Label3 
      Caption         =   "Value (mark at 90°)"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Value: (at 180°) _____"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Please SLOWLY hover the mouse over this rectangle width to test"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   6600
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Const vbPI = 3.141592654
Const Deg2Rad = vbPI / 180 'Degrees to Radians



'This routine is to demonstrate how to call the DrawCircularMeter function
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Caption = Str(X)
' Here I'm calling it with x,y,radius to cero, so it will be centered
Call DrawCircularMeter(Picture2, 0, 0, 0, 0, 100, X, 20, 90, True, vbRed, vbBlack, vbBlue)
' Here I'm calling it with a center off the limits of the picture2
Call DrawCircularMeter(Picture1, 500, 100, 550, 0, 200, X, 80, 180, False, vbWhite, vbBlue, vbYellow)
End Sub




'Parameters description
'PictureBox      : where to draw. MUST be declared in pixels
'Puntomediox     : X offset of DIAL center. of zeroed, it is calculated at the center of the picturebox
'Puntomedioy     : Y offset of DIAL center. of zeroed, it is calculated at the center of the picturebox
'Radio           : Radius of DIAL center. of zeroed, it is calculated from the center of the picturebox
'valormin        : Minium value to plot (normally zero)
'valormax        : Maximum value to plot (using rule of 3, valormax=360 degrees of the dial)
'valor           : Value to plot in the units of measure used for valormin and valormax
'numeromarcas    : Number of ticks of the dial. Each one alternatively will be long and short. text will be plotted for long ones.
'anguloindicador :Angle in degrees where the 0 will start
'clockwise       : if true, 0 starts and moves clockwise as value increases

Private Sub DrawCircularMeter(ByVal pic As PictureBox, ByVal puntomediox, ByVal puntomedioy, ByVal radio, ByVal valormin As Long, ByVal valormax As Long, ByVal valor As Long, ByVal numeromarcas As Long, ByVal AnguloIndicador As Long, ByVal ClockWise As Boolean, ByVal TextColor As ColorConstants, ByVal BigTickCOlor As ColorConstants, ByVal LowTickColor As ColorConstants)
Dim AnguloGlobal As Double
Dim Grueso As Boolean
Dim IntValorActual As Long

Dim ValorDx, ValorActual As Double
Dim I, AnguloMin, AnguloMax, AnguloDx, AnguloActual As Double
Dim RadioMenor, RadioMayor, RadioCentro, RadioFuente As Double
Dim LoopIni As Double
Dim LoopEnd As Double
Dim LoopDx As Double
Dim PuntoFinalx, PuntoFinalY As Double
Dim PuntoInicialx, PuntoInicialY As Double
Dim PuntoFuentex, PuntoFuenteY As Double
Dim Cwise As Double

pic.Cls 'clears dial


valor = Minimo(valor, valormax)
AnguloGlobal = valor * 360 / valormax
If puntomediox = 0 And puntomedioy = 0 Then  'if no X,Y,Radius parameters are passed, they are calculated at the center of
 radio = Minimo(pic.Width, pic.Height) / 2   'the picturebox. radius is set to the smallest side of the rectangle
 puntomediox = pic.Width / 2
 puntomedioy = pic.Height / 2
End If
pic.Circle (puntomediox, puntomedioy), radio * 0.1
RadioMenor = radio * 0.8   'lenght of short ticks
RadioMayor = radio * 0.9   'lenght of long ticks

RadioFuente = radio * 0.6  'Radius position of text

angulominimo = 0
angulomaximo = 360
ValorDx = (valormax - valormin) / numeromarcas

AnguloDx = (angulomaximo - angulominimo) / numeromarcas
Grueso = True

If ClockWise Then 'defines clockwise or counterclockwise drawing
   LoopIni = angulomaximo
   LoopEnd = angulominimo
   LoopDx = -AnguloDx
   ValorActual = valormin
   Cwise = 1
Else
   LoopIni = angulominimo
   LoopEnd = angulomaximo
   LoopDx = AnguloDx
   ValorActual = valormax
   Cwise = -1
   End If
   
'loop for drawing ticks and text
For I = LoopIni To LoopEnd Step LoopDx
    AnguloActual = (AnguloGlobal - AnguloIndicador + I) * Deg2Rad * Cwise
    PuntoFuentex = Cos((AnguloActual)) * (RadioFuente) + puntomediox  'font position x
    PuntoFuenteY = Sin((AnguloActual)) * (RadioFuente) + puntomedioy  'font position y
    'Grueso is a flag used to draw alternately large and small ticks
    IntValorActual = Int(ValorActual)
    If Grueso = True Then
       Grueso = False
       pic.ForeColor = TextColor
       'Picture2.Line (PuntoFuentex, PuntoFuenteY)-(PuntoFuentex + Picture2.TextWidth(ValorActual), PuntoFuenteY + Picture2.TextHeight(ValorActual)), vbRed, B
       pic.CurrentX = PuntoFuentex - (pic.TextWidth(IntValorActual))
       pic.CurrentY = PuntoFuenteY - (pic.TextHeight(IntValorActual) / 2)
       pic.Font.Name = "Arial narrow"
       pic.FontSize = radio * 0.07
       If I <> angulominimo Then pic.Print IntValorActual 'evita encimar el 0 con el valor final
       pic.DrawWidth = 5
       pic.ForeColor = vbWhite
       PuntoInicialx = Cos((AnguloActual)) * (RadioMenor * 0.9) + puntomediox 'hace que la raya gruesa sea mas larga que la delgada
       PuntoInicialY = Sin((AnguloActual)) * (RadioMenor * 0.9) + puntomedioy
       pic.ForeColor = LowTickColor
    Else
       Grueso = True
       pic.DrawWidth = 1
       pic.ForeColor = LowTickColor
       PuntoInicialx = Cos((AnguloActual)) * (RadioMenor) + puntomediox
       PuntoInicialY = Sin((AnguloActual)) * (RadioMenor) + puntomedioy
       pic.ForeColor = BigTickCOlor
    End If
    PuntoFinalx = Cos((AnguloActual)) * (RadioMayor) + puntomediox
    PuntoFinalY = Sin((AnguloActual)) * (RadioMayor) + puntomedioy
    pic.Line (PuntoInicialx, PuntoInicialY)-(PuntoFinalx, PuntoFinalY)
    ValorActual = ValorActual + ValorDx * Cwise
Next I

End Sub

Public Function Minimo(ByVal valor1 As Long, ByVal valor2 As Long) As Long
 Minimo = valor2
 If valor1 < valor2 Then Minimo = valor1
End Function


