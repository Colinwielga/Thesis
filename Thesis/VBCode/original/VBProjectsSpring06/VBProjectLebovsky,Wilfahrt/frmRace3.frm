VERSION 5.00
Begin VB.Form frmRace3 
   BackColor       =   &H00000000&
   Caption         =   "Race with the Buggy"
   ClientHeight    =   5775
   ClientLeft      =   2295
   ClientTop       =   2475
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   10485
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   3960
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000C&
      Caption         =   "Quit"
      Height          =   735
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H8000000C&
      Caption         =   "Reset"
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H8000000C&
      Caption         =   "Previous Page"
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.PictureBox picDel 
      Height          =   975
      Left            =   240
      Picture         =   "frmRace3.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H8000000C&
      Caption         =   "Start/Continue the Race!"
      Height          =   735
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.PictureBox picBuggy 
      Height          =   1695
      Left            =   240
      Picture         =   "frmRace3.frx":0B9D
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   600
      Width           =   2415
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
      Height          =   975
      Left            =   840
      TabIndex        =   8
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Finish"
      Height          =   375
      Left            =   9360
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmRace3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racing
'frmRace3(frmRace3.frm)
'Clay Wilfahrt and Andy Lebovsky
'3/22/06
'The purpose of this form is for the player to race the cars by answering questions which they should have learned the answers to earlier in the project.

Option Explicit
Dim Answer As String, status As Integer
'Returns you to previous screen
Private Sub cmdPrevious_Click()
    Close #1
    frmCar.Show
    frmRace3.Hide
    
End Sub
'Exits the program
Private Sub cmdquit_Click()
End
End Sub
'Resets the Cars
Private Sub cmdreset_Click()
    picBuggy.Left = 120
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
        picBuggy.Move (picBuggy.Left + 1700)
    Else
        MsgBox "You got a flat tire!", , "Patience"
        Timer1.Enabled = False
    End If
    If picBuggy.Left >= 7680 Then
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
