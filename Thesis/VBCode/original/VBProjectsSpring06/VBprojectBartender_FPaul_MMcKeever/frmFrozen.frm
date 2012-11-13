VERSION 5.00
Begin VB.Form frmFrozen 
   BackColor       =   &H80000012&
   Caption         =   "Frozen Drinks"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   Picture         =   "frmFrozen.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   315
      Left            =   7560
      TabIndex        =   5
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdFindDrink 
      Caption         =   "Display Drink Information"
      Height          =   735
      Left            =   7560
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00000000&
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   8160
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblFrozen10 
      BackStyle       =   0  'Transparent
      Caption         =   "10.Frostbite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblFrozen9 
      BackStyle       =   0  'Transparent
      Caption         =   "9.Pina Colada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblFrozen8 
      BackStyle       =   0  'Transparent
      Caption         =   "8.Cantaloupe Cup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblFrozen7 
      BackStyle       =   0  'Transparent
      Caption         =   "7.Frozen Mint Julep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblFrozen6 
      BackStyle       =   0  'Transparent
      Caption         =   "6.Frozen Matador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblFrozen5 
      BackStyle       =   0  'Transparent
      Caption         =   "5.Frozen Apple"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblFrozen4 
      BackStyle       =   0  'Transparent
      Caption         =   "4.Frozen Tequila Screwdriver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblFrozen3 
      BackStyle       =   0  'Transparent
      Caption         =   "3.Mudslide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblFrozen2 
      BackStyle       =   0  'Transparent
      Caption         =   "2.Margarita"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblFrozen1 
      BackStyle       =   0  'Transparent
      Caption         =   "1.Daiquiri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblInput 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Input Drink Number for Recipe and Drink Preparation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6120
      TabIndex        =   2
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblOutput1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   2400
      TabIndex        =   1
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label lblOutput 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   4680
      Width           =   3375
   End
End
Attribute VB_Name = "frmFrozen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Bartending School
'frmFrozen(Frozen Drinks)
'By Fred Paul & Michael McKeever
'March 22,2006
''The Frozen Drink form asks the user for a drink number via an
'Input box and displays the recipe in a label caption.

'Declare Variable
Dim Numbers(1 To 20) As Single
Dim Drink(1 To 20) As String
Dim Recipes(1 To 20) As String
Dim Pos, Size As Integer
Dim DrinkNumber As Integer


Private Sub cmdBack_Click()
'When clicked this button returns user to the Bartender page
'I also mkes the frozen drinks disappear
    frmFrozen.Hide
    frmBartender.Show
End Sub

Private Sub cmdFindDrink_Click()
    'This Button retrieves the Input Number and matches/Displays
    'it with it's corresponding Drink and recipes
    Dim found As Boolean
    found = False

    Pos = 0
    Open App.Path & "\frozen.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Numbers(Pos), Drink(Pos), Recipes(Pos)
    Loop
    Close #1
    Size = Pos
    Pos = 0
    Do While found = False And Pos < Size
        Pos = Pos + 1
        If DrinkNumber = Numbers(Pos) Then
            lblOutput1.Caption = Recipes(Pos)
            lblOutput.Caption = Drink(Pos)
            found = True
        End If
    Loop
    If found = False Then
        MsgBox "Drink Number Not Found", , "Error"
    End If
End Sub
Private Sub txtInput_Change()
    'This Declares drinknumber as an integer entered in the input
    'box
    DrinkNumber = txtInput.Text
End Sub

