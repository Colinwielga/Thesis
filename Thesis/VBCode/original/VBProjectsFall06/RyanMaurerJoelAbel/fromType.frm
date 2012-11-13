VERSION 5.00
Begin VB.Form Defuse 
   BackColor       =   &H000000FF&
   Caption         =   "Defuse"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1575
   End
   Begin VB.OptionButton cmdAgent20 
      BackColor       =   &H000000FF&
      Caption         =   "00 Agent"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.OptionButton cmdAgent30 
      BackColor       =   &H000000FF&
      Caption         =   "Secret Agent"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.OptionButton cmdAgent40 
      BackColor       =   &H000000FF&
      Caption         =   "Agent"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.OptionButton cmdAgent60 
      BackColor       =   &H000000FF&
      Caption         =   "Rookie Agent"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   360
      Top             =   120
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txt4 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txt3 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txt2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   2595
      Left            =   5040
      Picture         =   "fromType.frx":0000
      Top             =   2880
      Width           =   2340
   End
   Begin VB.Label bomb 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   9
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lbl4 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label lbl3 
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lbl2 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lbl1 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "Defuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Defuse by Ryan Maurer and Joel Abel 11/06



Private Sub cmdAbout_Click() 'displays a help box about the game
    MsgBox ("Select a difficulty setting and then type the correct numbers in the correct text boxes to defuse the bomb and save the day!")
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdStart_Click()
' This section contains the code for loading the timer
'and displaying the correct time in the display box
    Timer.Enabled = True
        lbl1.Caption = (Rnd) * 10 'this generates a random number under 10
        lbl2.Caption = (Rnd) * 100 'this generates a random number under 100
        lbl3.Caption = (Rnd) * 1000 ' this generates a random number under 1000
        lbl4.Caption = (Rnd) * 1 'this generates a random number under 1
    If cmdAgent60.Value = True Then 'this gives the display a value of 60
       bomb = "60"
    End If
    If cmdAgent40.Value = True Then 'this gives the display a value of 40
       bomb = "40"
    End If
    If cmdAgent30.Value = True Then ' this gives the display a value of 30
        bomb = "30"
    End If
    If cmdAgent20.Value = True Then 'this gives the display a value of 20
        bomb = "20"
    End If


End Sub

Private Sub Form_Load()
    Timer.Enabled = False ' The form loads with the timer disabled
End Sub

Private Sub Timer_Timer()
    If txt1.Text = lbl1.Caption And txt2.Text = lbl2.Caption And txt3.Text = lbl3.Caption And txt4.Text = lbl4.Caption Then
        Winner.Show
        Defuse.Hide
        Timer.Enabled = False 'this code loads the winning form if the person has correctly written all four numbers
    End If
    If cmdAgent60.Value = True Then 'these codes start the timer with the appropriate time and count it down second by second
        bomb.Caption = bomb
        bomb = bomb - 1
    End If
    If cmdAgent40.Value = True Then
        bomb.Caption = bomb
        bomb = bomb - 1
    End If
    If cmdAgent30.Value = True Then
       bomb.Caption = bomb
       bomb = bomb - 1
    End If
    If cmdAgent20.Value = True Then
        bomb.Caption = bomb
        bomb = bomb - 1
    End If
    If bomb.Caption <= 0 Then 'this code loads the loser form in the person failed to type the correct numbers
        Timer.Enabled = False
        Loser.Show
        Defuse.Hide
    End If
End Sub
