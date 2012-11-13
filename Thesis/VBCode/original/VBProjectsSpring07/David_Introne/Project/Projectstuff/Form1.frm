VERSION 5.00
Begin VB.Form FrmDog_Pound_Main 
   Caption         =   "Take a dog home and Orient it's owner!"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton HighScorez 
      BackColor       =   &H00FFC0C0&
      Caption         =   "High Score"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton CmdStart 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   8280
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton CmdAbout 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Citation"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label LblAbout 
      Caption         =   "Acknowledgements: Imads ""Bubble Sort Complete"" Program was modeled in form ""training"" under button ""Sort""  "
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pick an owner, Pick a pup and get started!"
      BeginProperty Font 
         Name            =   "PMingLiU"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "The Dog Pound "
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   7560
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10080
   End
End
Attribute VB_Name = "FrmDog_Pound_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAbout_Click(Index As Integer)

click = click + 1
'this sets a counter to 1
    Select Case click 'this lets you click up to 13 times
        Case 1 'every click is a case, every case is an action
        LblAbout.Visible = True
        Case 2
        LblAbout.Visible = False
        Case 3
        LblAbout.Visible = True
        Case 4
        LblAbout.Visible = False
        Case 5
        LblAbout.Visible = True
        Case 6
        LblAbout.Visible = False
        Case 7
        LblAbout.Visible = True
        Case 8
        LblAbout.Visible = False
        Case 9
        LblAbout.Visible = True
        Case 10
        LblAbout.Visible = False
        Case 11
        LblAbout.Visible = True
        Case 12
        LblAbout.Visible = False
        Case 13
        LblAbout.Visible = True
        End Select
End Sub





Private Sub CmdManual_Click()
    click2 = click2 + 1


    'this sets a counter to 1
    Select Case click2 'this lets you click up to 13 times
        Case 1 'every click is a case, every case is an action
        lblManual.Visible = True
        Case 2
        lblManual.Visible = False
        Case 3
        lblManual.Visible = True
        Case 4
        lblManual.Visible = False
        Case 5
        lblManual.Visible = True
        Case 6
        lblManual.Visible = False
        Case 7
        lblManual.Visible = True
        Case 8
        lblManual.Visible = False
        Case 9
        lblManual.Visible = True
        Case 10
        lblManual.Visible = False
        Case 11
        lblManual.Visible = True
        Case 12
        lblManual.Visible = False
        Case 13
        lblManual.Visible = True
        End Select
End Sub

Private Sub CmdStart_Click(Index As Integer)
ctr10 = 1 'begins game
FrmDog_Pound_Main.Hide
Player.Show
End Sub

Private Sub CmdQuit_Click()
BeginQuit.Show 'takes player to funny quit screen
End Sub

Private Sub HighScorez_Click()
If ctr10 = 1 Then 'this sets a counter to 1 when you click start, so you can't click high score
'MsgBox "The high score is " & Score & " held by " & HighScore
MsgBox HighScore & " Had the high score of " & Score
End If
End Sub


