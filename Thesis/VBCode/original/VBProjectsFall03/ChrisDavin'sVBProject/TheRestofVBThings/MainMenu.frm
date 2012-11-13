VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "Main Menu"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   2055
      Left            =   3480
      Picture         =   "Main Menu.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   2715
      TabIndex        =   4
      Top             =   480
      Width           =   2775
   End
   Begin VB.PictureBox Picture4 
      Height          =   1695
      Left            =   360
      Picture         =   "Main Menu.frx":30A0
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   1920
      Picture         =   "Main Menu.frx":4E5F
      ScaleHeight     =   1035
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   4080
      Picture         =   "Main Menu.frx":7FC8
      ScaleHeight     =   1755
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      Picture         =   "Main Menu.frx":9189
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Designed by Chris Davin"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "A Talk with Jasmine"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Whack a Goblin"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Pokemon Sort"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Project Quiz"
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Information"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Click on any picture to begin having an awsome time.  Click the Samurai in the upper left for more info."
      Height          =   1215
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to Chris's Crazy Project of Coolness."
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MemoryGamesEtc (Chris Davin's VB Project.vbp)
'Form Name : frmMainMenu (Main Menu.frm)
'Author: Chris Davin
'Date Written: October 29, 2003
'Purpose of Form: This purpose of this project as a whole
                'is to use my knowledge of VB
                'to create a fun program to play with.
                'The only problem the program solves is boredom.
                'This is of computational significance because
                'much programming in the world is games.
                'This specific form is to have a
                'Main Menu for the players
                'to explore the program from.
                'This is where you go between forms.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit

Private Sub Form_Load()

End Sub

'Click this to take you to the information form.
Private Sub Picture1_Click()
    frmMainMenu.Hide
    frmInformation.Show
End Sub
'This will take you to the Pokemon Sorting Form
Private Sub Picture2_Click()
    frmMainMenu.Hide
    frmPokemonSort.Show
End Sub
'This will take you to the Project Quiz
'however before you go you have to input yes
'doing this acts as a warning because you can't return from the Quiz.
'If you've already attempted the Quiz A returning message will appear instead.
Private Sub Picture3_Click()
    Dim A As String
    If Atempt = True Then
            MsgBox "Show me what you've learned.", , "Quiz Master"
            frmMainMenu.Hide
            frmQuiz.Show
        Else
            A = InputBox("This leads to the project Quiz.  It will quiz you about various things from the project.  You can't return from it.  Do you want to Take the Project Quiz Now?", "Yes or No")
            If A = "Yes" Then
                    frmMainMenu.Hide
                    frmQuiz.Show
                'If you change your mind message
                Else
                    MsgBox "I'll be waiting.", , "Quiz Master"
            End If
    End If
End Sub
'Takes you to the Jasmine form.
'If you've been and entered your name already
'Jasmine will say Welcome back
Private Sub Picture4_Click()
    If Visit = True Then
            MsgBox "Welcome back", , "Jasmine"
            frmMainMenu.Hide
            frmJasmine.Show
        Else
            frmMainMenu.Hide
            frmJasmine.Show
    End If
End Sub
'Takes you to the WhackaGoblin Form
Private Sub Picture5_Click()
    frmMainMenu.Hide
    frmWhackaGoblin.Show
End Sub
