VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Other Value"
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Recalc"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   3720
      Width           =   975
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "0"
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Minimum Calculated scrap = "
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "N° of strips = "
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Obtained  scrap ="
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Entradas() As String            'Entered elements Matrix
Dim cuenta As Byte                   'Counter
Dim ValorNom As Single             'Proposed nominal length
Dim NumElem As String             'N° elements in the matrix
Dim Descarte As Single              'Total scrap




Private Sub Command1_Click()
Dim Elegido As Byte                    'choiced element (indice)
Dim Media$                                 'choiced element value
Dim Parcial$                                'Variable for output list
Dim Suma As Single                    'lengths acumulator
Dim Num As Byte                        'Number of elements in list2
Dim contador As Byte                  'Counter
Dim vueltas As Integer
Dim MaxDesc As Single
Dim MinDesc As Single

Suma = 0
contador = 0
Descarte = 0
vueltas = 0
MaxDesc = 0.5
MinDesc = 1000
Num = NumElem
ProgressBar1.Min = 1
ProgressBar1.Max = 1000
ProgressBar1.Visible = True
ProgressBar1.Value = ProgressBar1.Min

Randomize

For vueltas = 1 To 1000
ProgressBar1.Value = vueltas

Arranque:
Elegido = Int((Num * Rnd) + 1)                         'Choice random element

List1.Selected(Elegido - 1) = True                    'Mark an element
Media = List1.List(Elegido - 1)                         'give value to Media
List1.RemoveItem (Elegido - 1)                        'erase it from de Lista1


    Suma = Suma + Val(Media)                               'Add to suma
    List2.AddItem Media                                   'Add a List2


If Suma > ValorNom Then                                         'if acummulated value is upper
Suma = Suma - Val(Media)                                             'than nominal value, then rest the
Else:                                                                       'last value, else pass it to
List2.RemoveItem (List2.ListCount - 1)                       'List2 and to Parcial
Parcial = Parcial + " " + Media
End If

Num = Num - 1

If List1.ListCount = 0 Then GoTo Segundo

GoTo Arranque


Segundo:
Descarte = Descarte + ValorNom - Suma

If List2.ListCount = 0 Then
    List3.AddItem Parcial
    'Command2.SetFocus
           GoTo Final
Else:
    List3.AddItem Parcial
           Parcial = ""
           Media = ""
           Suma = 0
End If

Num = List2.ListCount

For contador = 1 To Num
List2.Selected(contador - 1) = True
Media = List2.List(contador - 1)
List1.AddItem Media
Next contador

List2.Clear
Media = ""

GoTo Arranque

Final:
Label1.Caption = Str(Descarte)
Label2.Caption = List3.ListCount
Label6.Caption = Str(MinDesc)
If Descarte <= MinDesc Then MinDesc = Descarte
If Descarte <= MaxDesc Then Exit For
List2.Clear
List3.Clear
Suma = 0
contador = 0
Num = NumElem
Parcial = ""
Media = ""
Descarte = 0
IntoDatos

Next vueltas

ProgressBar1.Visible = False
ProgressBar1.Value = ProgressBar1.Min

End Sub

Private Sub Command2_Click()
List2.Clear
List3.Clear
Suma = 0
contador = 0
Num = 0
Parcial = ""
Media = ""
Descarte = 0
IntoDatos
End Sub

Private Sub Command3_Click()

Printer.Print "Largo nominal = " & ValorNom
Printer.Print "N° de items = " & NumElem
Printer.Print "N° de tiras = " & List3.ListCount
Printer.Print "Descarte = " & Str(Descarte)
Printer.Print

For cuenta = 1 To List3.ListCount
List3.Selected(cuenta - 1) = True
Media = List3.List(cuenta - 1)
Printer.Print Media
Next cuenta
Printer.EndDoc

End Sub

Private Sub Command4_Click()
ComeIn
End Sub

Private Sub Form_Load()

'ProgressBar1.Align = vbAlignBottom
'ProgressBar1.Visible = False

cuenta = 0

NumElem = InputBox("Start", "Enter N° of elements", "")

ReDim Entradas(NumElem)

ComeIn

For cuenta = 1 To NumElem
    Entradas(cuenta) = InputBox("Enter Value " & Str(cuenta), "Enter Data", "")
Next cuenta

IntoDatos

End Sub

Sub IntoDatos()

For cuenta = 1 To NumElem
       List1.AddItem Entradas(cuenta)
Next cuenta

End Sub

Sub ComeIn()
ValorNom = InputBox("Standard Value", "Enter nominal length", "")
End Sub



