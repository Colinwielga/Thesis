VERSION 5.00
Begin VB.Form frm10 
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   Picture         =   "frm10.frx":0000
   ScaleHeight     =   7320
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   2835
      TabIndex        =   6
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eat Lambis Bread"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "End Your Journey"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Leave Gollum"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Trust Gollum"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm10.frx":B63E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "frm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox ("Although Gollum is very shady and nothing would please you more then ending his life at this very moment you realize any short cut to Morodor would be most helpful and you gain a new member of your fellowship.")
    frm10.Hide
    frm11.Show
End Sub

Private Sub Command2_Click()
    frm10.Hide
    frm10Die.Show
End Sub

Private Sub Command3_Click()
    frm10.Hide
    frm9.Show
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub Command5_Click()
    If Lambis > 0 Then
        Lambis = Lambis - 1
        picResults.Cls
        picResults.Print "You have"; Lambis; " pieces of Lambis Bread."
    End If
End Sub
