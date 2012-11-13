VERSION 5.00
Begin VB.Form frmFinal 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFinal 
      Height          =   975
      Left            =   3600
      TabIndex        =   2
      Top             =   3720
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Enter Your answer here (jeopardy answer form):"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   6615
   End
   Begin VB.Label lblFinal 
      Caption         =   "Who is King Tut"
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Anthough his rule could be considered brief and insignificant, he is the most famous of all the pharohs of ancient egypt."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   6975
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'This button compares the user's answer to the correct answer.
    'If it is correct, their wager is added to their total.
    'If it is incorrect, the wager is subtracted from their total.
    'Finally, the Check form is displayed and the Final form is hidden.
    If txtFinal.Text = lblFinal.Caption Then
        Sum = Sum + Wager
        MsgBox ("Correct!")
    Else
        Sum = Sum - Wager
        MsgBox ("Wrong!")
        
    End If
    frmCheck.Show
    frmFinal.Hide
End Sub
