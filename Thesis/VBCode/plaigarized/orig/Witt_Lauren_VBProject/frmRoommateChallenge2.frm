VERSION 5.00
Begin VB.Form frmRoommateChallenge2 
   BackColor       =   &H00000000&
   Caption         =   "Roommate Challenge Log In"
   ClientHeight    =   9810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRCPlay 
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1080
      Picture         =   "frmRoommateChallenge2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7800
      Width           =   4815
   End
   Begin VB.CommandButton cmdplayer2login 
      BackColor       =   &H00FF00FF&
      Caption         =   "Player 2 Log in"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdPlayer4login 
      BackColor       =   &H00FF80FF&
      Caption         =   "Player 4 Log in "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdPlayer3login 
      BackColor       =   &H00FF00FF&
      Caption         =   "Player 3 Log in"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdplayer1login 
      BackColor       =   &H00FF80FF&
      Caption         =   "Player 1 Log in"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1575
   End
   Begin VB.PictureBox picKatie 
      Height          =   1575
      Left            =   2880
      Picture         =   "frmRoommateChallenge2.frx":151C2
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox picMe 
      Height          =   1695
      Left            =   4680
      Picture         =   "frmRoommateChallenge2.frx":1D7B4
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox picEmily 
      Height          =   1695
      Left            =   2880
      Picture         =   "frmRoommateChallenge2.frx":25556
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox picSteph 
      Height          =   1575
      Left            =   4680
      Picture         =   "frmRoommateChallenge2.frx":2FC90
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox picHeather 
      Height          =   1575
      Left            =   1080
      Picture         =   "frmRoommateChallenge2.frx":37C02
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox picCaitlin 
      Height          =   1695
      Left            =   1080
      Picture         =   "frmRoommateChallenge2.frx":40B24
      ScaleHeight     =   1635
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblRCLogin 
      BackColor       =   &H00FF80FF&
      Caption         =   "               LOG IN!"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   240
      Width           =   7695
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00000000&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00000000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00000000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00000000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00000000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3120
      Width           =   375
   End
End
Attribute VB_Name = "frmRoommateChallenge2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdplayer1login_Click()
RCPlayer1Name = InputBox("Enter Your Name", "Name")
RCPlayer1Pic = InputBox("Pick your picture by entering the number below your photo", "Picture")
If RCPlayer1Pic > 6 Or RCPlayer1Pic < 1 Then
    MsgBox "There is no picture matching that number", , "Error"
End If
End Sub

Private Sub cmdplayer2login_Click()
RCPlayer2Name = InputBox("Enter Your Name", "Name")
RCPlayer2Pic = InputBox("Pick your picture by entering the number below your photo", "Picture")
If RCPlayer2Pic > 6 Or RCPlayer2Pic < 1 Then
    MsgBox "There is no picture matching that number", , "Error"
End If
End Sub

Private Sub cmdPlayer3login_Click()
RCPlayer3Name = InputBox("Enter Your Name", "Name")
RCPlayer3Pic = InputBox("Pick your picture by entering the number below your photo", "Picture")
If RCPlayer3Pic > 6 Or RCPlayer3Pic < 1 Then
    MsgBox "There is no picture matching that number", , "Error"
End If
End Sub

Private Sub cmdPlayer4login_Click()
RCPlayer4Name = InputBox("Enter Your Name", "Name")
RCPlayer4Pic = InputBox("Pick your picture by entering the number below your photo", "Picture")
If RCPlayer4Pic > 6 Or RCPlayer4Pic < 1 Then
    MsgBox "There is no picture matching that number", , "Error"
End If
End Sub

Private Sub cmdRCPlay_Click()
    frmRoommateChallenge2.Hide
    frmRoommateChallenge3.Show
End Sub
