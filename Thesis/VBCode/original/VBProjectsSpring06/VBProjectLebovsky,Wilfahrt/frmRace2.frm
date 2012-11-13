VERSION 5.00
Begin VB.Form frmRace2 
   BackColor       =   &H00000000&
   Caption         =   "Race with the Classic"
   ClientHeight    =   5400
   ClientLeft      =   2400
   ClientTop       =   2910
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   10665
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7200
      Top             =   3720
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000C&
      Caption         =   "Quit"
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdreset 
      BackColor       =   &H8000000C&
      Caption         =   "Reset"
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdprevious 
      BackColor       =   &H8000000C&
      Caption         =   "Previous Page"
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.PictureBox picDel 
      Height          =   975
      Left            =   120
      Picture         =   "frmRace2.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H8000000C&
      Caption         =   "Start/Continue the Race!"
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox picClassic 
      Height          =   1095
      Left            =   120
      Picture         =   "frmRace2.frx":0B9D
      ScaleHeight     =   1035
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Clay Wilfahrt and Andy Lebovsky"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1080
      TabIndex        =   8
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Finish"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmRace2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racing
'frmRace2(frmRace2.frm)
'Clay Wilfahrt and Andy Lebovsky
'3/22/06
'The purpose of this form is for the player to race the cars by answering questions which they should have learned the answers to earlier in the project.

Option Explicit
Dim Answer As String, status As Integer
'Brings you to previous screen
Private Sub cmdPrevious_Click()
    Close #1
    frmCar.Show
    frmRace2.Hide
    
End Sub
'Exits the program
Private Sub cmdquit_Click()
End
End Sub
'Resets the Cars
Private Sub cmdreset_Click()
    picClassic.Left = 120
    picDel.Left = 120
    Close #1
    Open App.Path & "\QA.txt" For Input As #1
    
End Sub
'Starts the Timer
Private Sub cmdStart_Click()
    Timer1.Enabled = True
    
   
End Sub


Private Sub Form_Load()
    Timer1.Enabled = False
    Open App.Path & "\QA.txt" For Input As #1
End Sub

'Races the cars across the screen
Private Sub Timer1_Timer()
    If Not EOF(1) Then
        Input #1, Q, a
    End If
    Answer = InputBox(Q, "Question")
    If Answer = a Then
        picClassic.Move (picClassic.Left + 1700)
    Else
        MsgBox "You got a flat tire!", , "Patience"
        Timer1.Enabled = False
    End If
    If picClassic.Left >= 7680 Then
        MsgBox "You DID NOT wreck your car! You Win", , "Winner"
        Timer1.Enabled = False
        status = 1
    Close #1
    End If
    
    picDel.Move (picDel.Left + 1400)
    If picDel.Left > 7680 And status = 0 Then
        MsgBox "You lost the race!", , "Loser"
        Timer1.Enabled = False
    Close #1
    End If
End Sub

