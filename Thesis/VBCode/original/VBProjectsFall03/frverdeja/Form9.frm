VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00000000&
   Caption         =   "Form9"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9240
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdNext7 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "~frank verdeja"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"Form9.frx":D139
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
      Height          =   3735
      Left            =   9120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the seventh form in the lesson, and allows user to continue or quit.


Private Sub cmdNext7_Click()
Form10.Show
Form9.Hide

End Sub

Private Sub Command1_Click()
Form1.Show
Form9.Hide

End Sub

