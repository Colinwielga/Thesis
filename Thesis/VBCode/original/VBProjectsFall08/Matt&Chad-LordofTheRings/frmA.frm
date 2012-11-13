VERSION 5.00
Begin VB.Form frmA 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   Picture         =   "frmA.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Leave Them To Fend For Themselves"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stay And Defend the People of Rohan At Helm's Deep"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End Your Journey"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   $"frmA.frx":7CC8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   8175
   End
End
Attribute VB_Name = "frmA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmA.Hide
    frm8.Show
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    MsgBox ("You decide to escort the villagers to Helm's Deep and although this will slow you down you understand the importance of making sure everyone you can help receives it.  You have made a valiant decision and the battle at Helm's Deep is won because of your expertise leadership and you have saved more lives then you will ever know.  Although Saroun's army isn't defeated they suffered a blow they weren't expecting.")
    frmA.Hide
    frmB.Show
End Sub

Private Sub Command4_Click()
    frmA.Hide
    frmaDie.Show
End Sub
