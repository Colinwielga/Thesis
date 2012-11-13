VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   Caption         =   "Form6"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Onyx"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   5175
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn1 
      BackColor       =   &H00000000&
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton cmdNext1 
      BackColor       =   &H00000000&
      Caption         =   "Next"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   2295
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
      Left            =   9120
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"Form6.frx":6E21
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
      Height          =   4215
      Left            =   7800
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the fourth form in the lesson, and allows user to continue or quit.

Private Sub cmdNext1_Click()
Form7.Show
Form6.Hide

End Sub

Private Sub cmdReturn1_Click()
Form1.Show
Form6.Hide

End Sub
