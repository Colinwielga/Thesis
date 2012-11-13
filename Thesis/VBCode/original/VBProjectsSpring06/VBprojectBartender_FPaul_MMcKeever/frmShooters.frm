VERSION 5.00
Begin VB.Form frmShooters 
   Caption         =   "Shooters"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   Picture         =   "frmShooters.frx":0000
   ScaleHeight     =   5955
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back "
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindDrink 
      Caption         =   "Display Drink Information"
      Height          =   735
      Left            =   7680
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label lblShooters15 
      BackStyle       =   0  'Transparent
      Caption         =   "14.Woo Woo"
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
      Left            =   0
      TabIndex        =   21
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblShooters14 
      BackStyle       =   0  'Transparent
      Caption         =   "13.Windex"
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
      Left            =   0
      TabIndex        =   20
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblShooters13 
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
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblShooters12 
      BackStyle       =   0  'Transparent
      Caption         =   "12.Tequila Shot"
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
      Left            =   0
      TabIndex        =   18
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblShooters11 
      BackStyle       =   0  'Transparent
      Caption         =   "11.Red Death"
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
      Left            =   0
      TabIndex        =   17
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblShooters10 
      BackStyle       =   0  'Transparent
      Caption         =   "10.Orgasm"
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
      Left            =   0
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblShooters9 
      BackStyle       =   0  'Transparent
      Caption         =   "9.Mind Eraser"
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
      Left            =   0
      TabIndex        =   15
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblShooters8 
      BackStyle       =   0  'Transparent
      Caption         =   "8.Liquid Cocaine"
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
      Left            =   0
      TabIndex        =   14
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblShooters7 
      BackStyle       =   0  'Transparent
      Caption         =   "7.Jell-O Shots"
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
      Left            =   0
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblShooters6 
      BackStyle       =   0  'Transparent
      Caption         =   "6.Hurricane"
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
      Left            =   0
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblShooters5 
      BackStyle       =   0  'Transparent
      Caption         =   "5.Fireball"
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
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblShooters4 
      BackStyle       =   0  'Transparent
      Caption         =   "4.Dr. Pepper"
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
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblShooters3 
      BackStyle       =   0  'Transparent
      Caption         =   "3.Cough Syrup"
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
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblShooter2 
      BackStyle       =   0  'Transparent
      Caption         =   "2.Blow Job"
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
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   600
      TabIndex        =   7
      Top             =   240
      Width           =   15
   End
   Begin VB.Label lblShooters1 
      BackStyle       =   0  'Transparent
      Caption         =   "1.Anti-Freeze"
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
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lblInput 
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
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   6240
      TabIndex        =   2
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label lblOutput1 
      BackColor       =   &H80000012&
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
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   5400
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblOutput 
      BackColor       =   &H80000012&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmShooters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Bartending School
'frmShooters(Shooters)
'By Fred Paul & Michael McKeever
'March 22,2006
''The Shooters form asks the user for a drink number via an
'Input box and displays the recipe in a label caption.

'declare variables
    Dim Numbers(1 To 20) As Single
    Dim Drink(1 To 20) As String
    Dim Recipes(1 To 20) As String
    Dim Pos, Size As Integer
    Dim DrinkNumber As Integer


Private Sub cmdBack_Click()
    'This button hides the shooters form and returns the user to
    'the Bartender form.
    frmShooters.Hide
    frmBartender.Show
End Sub

Private Sub cmdFindDrink_Click()
     'This button reads the number from the input box and matches
    'it with its correlating drink and recipes from mixed.txt and
    'displays it in a lbl.caption
    Dim found As Boolean
    found = False

    Pos = 0
    Open App.Path & "\shooters.txt" For Input As #1
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


