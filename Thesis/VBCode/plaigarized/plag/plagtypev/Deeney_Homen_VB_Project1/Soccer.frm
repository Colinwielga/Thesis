VERSION 5.00
Begin VB.Form StartUp
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   Picture         =   "Soccer.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      Height          =   495
      Left            =   11880
      TabIndex        =   6
      Top             =   7680
      Width           =   495
   End
   Begin VB.CommandButton cmdEnterSite
      BackColor       =   &H0080FF80&
      Caption         =   "Click here to take a closer look at the players who shocked the world!"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton cmdShowPicture
      BackColor       =   &H0080FF80&
      Caption         =   "2006 inspired champions: who will wear the crown this time around?"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   3015
   End
   Begin VB.PictureBox picResults
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   600
      ScaleHeight     =   3915
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   960
      Width           =   7095
   End
   Begin VB.Label lblPlayMusic
      BackColor       =   &H00000000&
      Caption         =   "Double Click Image to Play Music! ==========================>"
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   5160
      Width           =   2655
   End
   Begin VB.OLE OLE1
      BackColor       =   &H00000000&
      Class           =   "Package"
      Height          =   735
      Left            =   4800
      OleObjectBlob   =   "Soccer.frx":38182
      SourceDoc       =   "M:\CS130\Project\01 Kernkraft 400.mp3"
      TabIndex        =   4
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label lblWelcome
      BackColor       =   &H00000000&
      Caption         =   "World Cup Soccer!"
      BeginProperty Font
         Name            =   "Garamond"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "StartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: WorldCup
'Form Name: StartUp
'Author: Brian Deeney and Nick Homen
'Date written: 2-20-10
'Objective: The purpose of this form is to Welcome the user to the program.  This form provides interface for music, picture, and access to the next form.  The overall purpose of this project is to provide fun and informative interface to create excitement and understanding for the World Cup.

Dim abcde as Long

'Allows access to the next form
Private Sub cmdEnterSite_Click()
StartUp.Hide
Stats.Show
End Sub

'Quit
Private Sub cmdQuit_Click()
End
End Sub
'Displays picture of previous champion to capture user's attention and interest
Private Sub cmdShowPicture_Click()
picResults.Cls
picResults.Picture = LoadPicture(App.Path & "\world_cup_2006_1_1600x1200(edited).jpg") 'when user presses button, the loaded image will appear in picture box
cmdEnterSite.Enabled = True
End Sub

Private Sub Form_Load() 'This prevents the main page from appearing before the user enters their name (or any other string)
Stats.Hide
Jerseys.Hide
Dim Found As Boolean
Found = False
    
'Provides a nice welcome
Do While Found = False
names = InputBox("The World Cup beckons: Please enter your name", "Name")
    If Len(names) = 0 Then  'makes sure that the user enters something
        MsgBox "Please enter your name", , "You forgot to submit a name"
    Else
        MsgBox names & ", this Summer, teams from across the globe will gear up to compete in South Africa for the 2010 World Cup; will you be ready?", , "Welcome!"
        Found = True
End If
Loop
 StartUp.Show 'the desired form will show up now
End Sub

