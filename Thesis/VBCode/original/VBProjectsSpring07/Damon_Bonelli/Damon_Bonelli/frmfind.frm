VERSION 5.00
Begin VB.Form frmfind 
   BackColor       =   &H000000FF&
   Caption         =   "Find player"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picoutput 
      Height          =   3735
      Left            =   480
      ScaleHeight     =   3675
      ScaleWidth      =   6675
      TabIndex        =   5
      Top             =   1680
      Width           =   6735
   End
   Begin VB.CommandButton cmdstatsfrm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Look up statistics"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdfindplayer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find player"
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtfindplayer 
      Height          =   975
      Left            =   6480
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdhome1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Home"
      Height          =   975
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblname 
      BackColor       =   &H000000FF&
      Caption         =   "Enter player name here (First name only):"
      Height          =   975
      Left            =   4440
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
'quits the program
End
End Sub

Private Sub cmdfindplayer_Click()
'this subroutine takes the input from the textbox on the screen
'and tests it against the names array to see if it matches any
'of the names in the file.  If it does, it will print the name,
'jersey number and position of the player in the picture box.
'If it does not, it will return an error message
picoutput.Cls
Dim found As Boolean, search As String, pos As Single
search = txtfindplayer.Text
Do While pos < ctr And found = False
    pos = pos + 1
    If LCase(search) = LCase(names(pos)) Then
        found = True
        picoutput.Print names(pos); " is number "; jersey(pos); " and plays "; position(pos)
    End If
Loop
If found = False Then
    MsgBox search & " is not a valid entry (Maybe you forgot to load the file).", , "Error"
End If
End Sub


Private Sub cmdhome1_Click()
'goes to the home screen
frmintro.Show
frmfind.Hide
End Sub

Private Sub cmdstatsfrm_Click()
'goes to the statistics screen
frmStats.Show
frmfind.Hide
End Sub


