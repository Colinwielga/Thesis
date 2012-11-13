VERSION 5.00
Begin VB.Form frmRace1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Race with the BMW"
   ClientHeight    =   5100
   ClientLeft      =   2610
   ClientTop       =   3210
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   10755
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7800
      Top             =   3480
   End
   Begin VB.PictureBox picDel 
      Height          =   975
      Left            =   120
      Picture         =   "frmRace1.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H8000000C&
      Caption         =   "Quit"
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdreset 
      BackColor       =   &H8000000C&
      Caption         =   "Reset"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H8000000C&
      Caption         =   "Previous Page"
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H8000000C&
      Caption         =   "Start/Continue the Race!"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox picBMW 
      Height          =   975
      Left            =   120
      Picture         =   "frmRace1.frx":0B9D
      ScaleHeight     =   915
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
      Left            =   960
      TabIndex        =   8
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Finish"
      Height          =   375
      Left            =   9120
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
Attribute VB_Name = "frmRace1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racing
'frmRace1(frmRace1.frm)
'Clay Wilfahrt and Andy Lebovsky
'3/22/06
'The purpose of this form is for the player to race the cars by answering questions which they should have learned the answers to earlier in the project.

Option Explicit
Dim status As Integer
'Returns to car screen
Private Sub cmdPrevious_Click()
    Close #1
    frmCar.Show
    frmRace1.Hide
    
End Sub
'Exits the program
Private Sub cmdquit_Click()
    End
End Sub
'Resets the cars
Private Sub cmdreset_Click()
    picBMW.Left = 120
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

'Races cars across the screen
Private Sub Timer1_Timer()
    If Not EOF(1) Then
        Input #1, Q, a
    End If
    Answer = InputBox(Q, "Question")
    If Answer = a Then
        picBMW.Move (picBMW.Left + 1700)
    Else
        MsgBox "You got a flat tire!", , "Patience"
        Timer1.Enabled = False
    End If
    If picBMW.Left >= 7680 Then
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




