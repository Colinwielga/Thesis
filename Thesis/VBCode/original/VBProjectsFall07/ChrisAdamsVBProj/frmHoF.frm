VERSION 5.00
Begin VB.Form frmHoF 
   Caption         =   "Congratulations, You have been selected to the Hall of Fame"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   Picture         =   "frmHoF.frx":0000
   ScaleHeight     =   4950
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSpeech 
      Caption         =   "Speech"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin VB.PictureBox picSpeech 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   360
      ScaleHeight     =   1395
      ScaleWidth      =   7515
      TabIndex        =   2
      Top             =   2520
      Width           =   7575
   End
   Begin VB.PictureBox picJersey 
      Height          =   2295
      Left            =   4920
      Picture         =   "frmHoF.frx":10B14
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox picBettman 
      Height          =   2295
      Left            =   360
      Picture         =   "frmHoF.frx":12994
      ScaleHeight     =   2235
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.OLE oleAnthem 
      AutoActivate    =   3  'Automatic
      BackColor       =   &H00C0FFFF&
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   855
      Left            =   3360
      OleObjectBlob   =   "frmHoF.frx":15422
      SourceDoc       =   "M:\CS130\ChrisAdamsVBProj\music\02 The State of Hockey.mp3"
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
End
Attribute VB_Name = "frmHoF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Quest for The Cup~Minnesota Wild Trivia Game

'Author: Chris Adams

'Date: November 2007

'This form is shown if the user makes it into the Hall of Fame by answering all questions correctly in certain levels

Private Sub cmdQuit_Click()
    
    'Show the Sources form
    frmHoF.Hide
    frmSources.Show

End Sub

Private Sub cmdSpeech_Click()
    
    'Loads the commissioner's speech into the picbox
    picSpeech.Print "Click on the Wild Logo Below to hear the Wild Anthem."
    picSpeech.Print " "
    picSpeech.Print "NHL Commissioner Gary Bettman:"
    picSpeech.Print "Welcome. Today we celebrate a successful journey to NHL supremecy."
    picSpeech.Print "A journay that included an All Star appearance, a Stanely Cup trophy, and a Conn Smythe Trophy."
    picSpeech.Print "So with no futher a due, I would like to welcome our newest member of the Hockey Hall of Fame,"
    picSpeech.Print "The captain of the Minnesota Wild, number "; Jersey; " "; Pos; " "; PlayerFirst; " "; PlayerLast
    cmdQuit.Enabled = True

End Sub
