VERSION 5.00
Begin VB.Form frmIntroduction 
   Caption         =   "Introduction"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExplore 
      BackColor       =   &H8000000A&
      Caption         =   "Click here to find out more about this wonderful profession!"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Left            =   10200
      MaskColor       =   &H80000002&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9480
      Width           =   4932
   End
   Begin VB.Label lblIntroduction 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Have you ever considered a career in Accounting???"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   9612
   End
   Begin VB.Image Image1 
      Height          =   16335
      Left            =   0
      Picture         =   "AccountingProject.frx":0000
      Top             =   0
      Width           =   19935
   End
End
Attribute VB_Name = "frmIntroduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Accounting Project
'Introduction Form
'Tony McLean
'3.31.2008
'The purpose of this form is to introduce the user to the program
Option Explicit
Private Sub cmdExplore_Click()
    'The following code is used to show and hide certain forms
    'when a command button is selected by the user
    frmIntroduction.Hide
    frmContents.Show
End Sub
Private Sub Form_Load()
    Dim Name As String
    'Allows the user to enter his or her name into the program
    Name = InputBox("Hello, what is your name?", "Enter Your Name")
    'Greets the user and introduces him/her to the program
    MsgBox "Hello " & Name & " I hope you enjoy my program!"
End Sub
