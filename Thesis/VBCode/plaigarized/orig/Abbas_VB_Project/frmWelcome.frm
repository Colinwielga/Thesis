VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   12255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   Picture         =   "frmWelcome.frx":0000
   ScaleHeight     =   12255
   ScaleWidth      =   14835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H0000C0C0&
      Caption         =   "Play The Intro Video!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   13320
      Picture         =   "frmWelcome.frx":92BE
      ScaleHeight     =   1935
      ScaleWidth      =   1455
      TabIndex        =   9
      Top             =   10320
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      Picture         =   "frmWelcome.frx":12070
      ScaleHeight     =   1935
      ScaleWidth      =   1575
      TabIndex        =   8
      Top             =   10320
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      Picture         =   "frmWelcome.frx":1B0D6
      ScaleHeight     =   2055
      ScaleWidth      =   1695
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   13320
      Picture         =   "frmWelcome.frx":24394
      ScaleHeight     =   1815
      ScaleWidth      =   1455
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdRecent 
      BackColor       =   &H0000C0C0&
      Caption         =   "Explore The 10 Most Recent Winners"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10200
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000C0C0&
      Caption         =   "Leave The Program!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   11160
      Width           =   2295
   End
   Begin VB.CommandButton cmdWhereNow 
      BackColor       =   &H0000C0C0&
      Caption         =   "Find Where They Are Now!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdWinners 
      BackColor       =   &H0000C0C0&
      Caption         =   "Explore Past Winners"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton cmdHistory 
      BackColor       =   &H0000C0C0&
      Caption         =   "History of the Trophy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   3120
      Picture         =   "frmWelcome.frx":2D20E
      ScaleHeight     =   7575
      ScaleWidth      =   8535
      TabIndex        =   5
      Top             =   4920
      Width           =   8535
   End
   Begin VB.Image imgheisman 
      Height          =   7875
      Left            =   2520
      Picture         =   "frmWelcome.frx":68E53
      Top             =   0
      Width           =   12600
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Heisman Trophy
'frmWelcome
'Kevin Abbas
'2-22-10
'Objective of form -  To introduce the subject of my project - The Heisman Trophy
'Overall Objective is to educate people about the history of the Heisman Trophy as well as provide information regarding who won the trophy and provide some interesting infromation about where each winner is today
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Dim filename As String, retVal As Long


Private Sub cmdExit_Click() 'exit the program and thank the user
    MsgBox ("Hope you enjoyed learning about the Heisman, have a nice day!")
    End
End Sub

Private Sub cmdHistory_Click() 'go to the history form and hide all others
    frmWelcome.Hide
    frmHistory.Show
    frmWinners.Hide
    frmWhereNow.Hide
    frmRecent.Hide
End Sub

Private Sub cmdPlay_Click()
    filename = "N:\HandIns\CS130_Herzfeld\Abbas_VB_Project\Heisman2.avi"
        filename = Chr(34) & filename & Chr(34)
        retVal = mciSendString("seek movie to start", 0, 0, 0)
        retVal = mciSendString("open " & filename & " type mpegvideo alias movie", 0, 0, 0)
        retVal = mciSendString("play movie", 0, 0, 0)
        
    Refresh
    

End Sub

Private Sub cmdRecent_Click() 'go to the Recent form and hide all others
    frmWelcome.Hide
    frmHistory.Hide
    frmWinners.Hide
    frmWhereNow.Hide
    frmRecent.Show
End Sub

Private Sub cmdWhereNow_Click() 'go to the WhereNow form and hide all others
    frmWhereNow.Show
    frmWelcome.Hide
    frmHistory.Hide
    frmWinners.Hide
    frmRecent.Hide
    
End Sub


Private Sub cmdWinners_Click() 'go to the Winners form and hide all others
    frmWinners.Show
    frmWelcome.Hide
    frmHistory.Hide
    frmWhereNow.Hide
    frmRecent.Hide
End Sub

Private Sub Form_Load()
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2


End Sub
