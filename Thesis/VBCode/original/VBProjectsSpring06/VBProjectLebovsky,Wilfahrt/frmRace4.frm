VERSION 5.00
Begin VB.Form frmRace4 
   BackColor       =   &H00000000&
   Caption         =   "Race with the Turbo Mini-Van"
   ClientHeight    =   6525
   ClientLeft      =   2730
   ClientTop       =   2160
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   10515
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6840
      Top             =   4560
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000C&
      Caption         =   "Quit"
      Height          =   735
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H8000000C&
      Caption         =   "Reset"
      Height          =   735
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H8000000C&
      Caption         =   "Previous Page"
      Height          =   735
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox picDel 
      Height          =   975
      Left            =   240
      Picture         =   "frmRace4.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H8000000C&
      Caption         =   "Start/Continue the Race!"
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox picVan 
      Height          =   2415
      Left            =   240
      Picture         =   "frmRace4.frx":0B9D
      ScaleHeight     =   2355
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   480
      Width           =   2295
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
      Left            =   960
      TabIndex        =   8
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Finish"
      Height          =   375
      Left            =   9480
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmRace4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racing
'frmRace4(frmRace4.frm)
'Clay Wilfahrt and Andy Lebovsky
'3/22/06
'The purpose of this form is for the player to race the cars by answering questions which they should have learned the answers to earlier in the project.

Option Explicit
Dim Answer As String, status As Integer
'Returns you to previous screen
Private Sub cmdPrevious_Click()
    Close #1
    frmCar.Show
    frmRace4.Hide
    
End Sub
'Exits the program
Private Sub cmdquit_Click()
End
End Sub
'Resets the cars
Private Sub cmdreset_Click()
    picVan.Left = 120
    picDel.Left = 120
    Close #1
    Open App.Path & "\QA.txt" For Input As #1
    
End Sub
'Starts the timer
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
        picVan.Move (picVan.Left + 1700)
    Else
        MsgBox "You got a flat tire!", , "Patience"
        Timer1.Enabled = False
    End If
    If picVan.Left >= 7680 Then
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
