VERSION 5.00
Begin VB.Form frmOpeningPage 
   BackColor       =   &H80000013&
   Caption         =   "Music by Michael!!! Opening Page"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   ForeColor       =   &H80000006&
   LinkTopic       =   "Form1"
   Picture         =   "frmOpeningPage.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMichael 
      AutoSize        =   -1  'True
      Height          =   5130
      Left            =   2280
      Picture         =   "frmOpeningPage.frx":1D7B6A
      ScaleHeight     =   5070
      ScaleWidth      =   5970
      TabIndex        =   3
      Top             =   960
      Width           =   6030
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Click Here to Quit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1560
      TabIndex        =   2
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H80000000&
      Caption         =   "Click Here to Begin ""Adventures In Music!!!"""
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6840
      TabIndex        =   1
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   $"frmOpeningPage.frx":1DAD27
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   4
      Top             =   6240
      Width           =   7575
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Welcome to Adventures In Music!!!"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmOpeningPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The intent of this program is to educate students of the age of 4th grade and higher on basic concepts of music and music theory, and some composers of music
'This program uses many different forms to educate and to assess the students ability to understand music
'This form is made to welcome the user, asking for there name which is then saved and used later on in the program

Private Sub cmdQuit_Click()                             'The purpose of this button is end the program
    MsgBox "Have a Great Day!", , "Have a Great Day!"   'Once they have clicked the button a message box pops up to give them one last farewell
End
End Sub

Private Sub cmdStart_Click()

NameGiven = InputBox("Please Enter Your Name Here", "Enter Your Name")      'An input box pops up to ask the user's name which is then saved in the public variable NameGiven
    frmOpeningPage.Visible = False                                          'The form frmOpeningPage is hidden
    frmLessonMainPage.Visible = True                                        'The form frmLessonMainPage is visible and ready for use
End Sub
