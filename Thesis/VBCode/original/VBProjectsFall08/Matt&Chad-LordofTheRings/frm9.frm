VERSION 5.00
Begin VB.Form frm9 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   Picture         =   "frm9.frx":0000
   ScaleHeight     =   7680
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "End Your Journey"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run And Hide"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stay And Fight"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm9.frx":11CE3
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   8415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm9.frx":11EF8
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frm9.Hide
    frm9Die.Show
End Sub

Private Sub Command2_Click()
    MsgBox ("You make the decision to hide and although you know the Ringwraiths will pursue you to the ends of the earth to capture the ring you have outsmarted them this time and live to venture another day.  You continue to Morodor, thankful for the wisdom your fellowship has shown so far.")
    frm9.Hide
    frm10.Show
End Sub

Private Sub Command3_Click()
    frm9.Hide
    frm8.Show
End Sub

Private Sub Command4_Click()
    End
End Sub
