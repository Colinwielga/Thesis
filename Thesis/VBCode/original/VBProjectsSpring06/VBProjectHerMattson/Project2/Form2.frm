VERSION 5.00
Begin VB.Form frmHerMattson2 
   BackColor       =   &H00FFFF00&
   Caption         =   "Personal Information"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form2"
   ScaleHeight     =   4785
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFName 
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtScore 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H008080FF&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H0080FF80&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtGender 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtAge 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Score on Quiz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblGender 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gender (m/f)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblAge 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmHerMattson2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson2
'Ee Her and Jennifer Mattson
'Written 3/14/06
'This menu allows user to input personal information and will export to a personal text file.


Private Sub cmdMainMenu_Click()
    frmHerMattson2.Hide
    frmHerMattson1.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSave_Click()
    Dim Pos As Integer
    Dim firstname, lastname, gender As String
    Dim Score, age As Integer
    Dim Userinput As Boolean
    firstname = txtFName.Text
    lastname = txtName.Text
    age = txtAge.Text
    gender = txtGender.Text
    Score = Val(txtScore.Text)
    Open App.Path & "\personal Information.txt" For Append As #1
    Write #1, firstname, lastname, age, gender, Score
    Close #1
End Sub


 
