VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   Caption         =   "Form5"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form6"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   5370
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdNext4 
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
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   2535
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
      Left            =   7200
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"Form5.frx":7853
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
      Height          =   3975
      Left            =   6240
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the third form in the lesson, and allows user to continue or quit.

Private Sub cmdNext4_Click()
Form6.Show
Form5.Hide

End Sub

Private Sub Command2_Click()
Form1.Show
Form5.Hide

End Sub
