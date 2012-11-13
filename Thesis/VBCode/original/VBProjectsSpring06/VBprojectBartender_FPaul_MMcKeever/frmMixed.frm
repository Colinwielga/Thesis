VERSION 5.00
Begin VB.Form frmMixed 
   Caption         =   "Mixed Drinks"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   Picture         =   "frmMixed.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00404080&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindDrink 
      BackColor       =   &H00404080&
      Caption         =   "Display Drink Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblMixed20 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "20.White Russian"
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
      Left            =   3600
      TabIndex        =   25
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblMixed19 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "19.Tom Collins"
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
      Left            =   3600
      TabIndex        =   24
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblMixed18 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "18.Tequila Sunrise"
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
      Left            =   3600
      TabIndex        =   23
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblMixed17 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "17.Sex on the Beach"
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
      Left            =   3600
      TabIndex        =   22
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblMixed16 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "16.Screwdriver"
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
      Left            =   3600
      TabIndex        =   21
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblMixed15 
      BackStyle       =   0  'Transparent
      Caption         =   "15.Rusty Nail"
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
      Left            =   1920
      TabIndex        =   20
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblMixed14 
      BackStyle       =   0  'Transparent
      Caption         =   "14.Old-Fashion"
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
      Left            =   1920
      TabIndex        =   19
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblMixed13 
      BackStyle       =   0  'Transparent
      Caption         =   "13.Martini"
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
      Left            =   1920
      TabIndex        =   18
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblMixed12 
      BackStyle       =   0  'Transparent
      Caption         =   "12.Manhatten"
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
      Left            =   1920
      TabIndex        =   17
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblMixed11 
      BackStyle       =   0  'Transparent
      Caption         =   "11.Mai Tai"
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
      Left            =   1920
      TabIndex        =   16
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblMixed10 
      BackStyle       =   0  'Transparent
      Caption         =   "10.Long Island"
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
      Left            =   1920
      TabIndex        =   15
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblMixed9 
      BackStyle       =   0  'Transparent
      Caption         =   "9.Kamikaze"
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
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblMixed8 
      BackStyle       =   0  'Transparent
      Caption         =   "8.Irish Coffee"
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
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblmixed7 
      BackStyle       =   0  'Transparent
      Caption         =   "7.Grasshopper"
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
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblmixed6 
      BackStyle       =   0  'Transparent
      Caption         =   "6.Gimlet"
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
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblMixed5 
      BackStyle       =   0  'Transparent
      Caption         =   "5.Fuzzy Navel"
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
      Top             =   960
      Width           =   1455
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
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label lblOutput 
      BackColor       =   &H80000007&
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
      Left            =   3600
      TabIndex        =   8
      Top             =   3720
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInput 
      BackStyle       =   0  'Transparent
      Caption         =   "Input Drink Number for Recipe and Preparation Instructions!"
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
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblMixed4 
      BackStyle       =   0  'Transparent
      Caption         =   "4.Cosmopolitan"
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
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblmixed3 
      BackStyle       =   0  'Transparent
      Caption         =   "3.Bloody Mary"
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
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblMixed2 
      BackStyle       =   0  'Transparent
      Caption         =   "2.Black Russian"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblMixed1 
      BackStyle       =   0  'Transparent
      Caption         =   "1.Amaretto Sour"
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
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmMixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Bartending School
'frmMixed(Mixed Drinks)
'By Fred Paul & Michael McKeever
'March 22,2006
'The Mixed Drink form asks the user for a drink number via an
'Input box and displays the recipe in a label caption.

'Declare variables
    Dim Numbers(1 To 20) As Single
    Dim Drink(1 To 20) As String
    Dim Recipes(1 To 20) As String
    Dim Pos, Size As Integer
    Dim DrinkNumber As Integer


Private Sub cmdBack_Click()
'This button hides the mixed form and returns you to the
'bartender form.
    frmMixed.Hide
    frmBartender.Show
End Sub

Private Sub cmdFindDrink_Click()
    'This button reads the number from the input box and matches
    'it with its correlating drink and recipes from mixed.txt and
    'displays it in a lbl.caption
    Dim found As Boolean
    found = False

    Pos = 0
    Open App.Path & "\mixed.txt" For Input As #1
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
            lblOutput.Caption = Recipes(Pos)
            lblOutput1.Caption = Drink(Pos)
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
