VERSION 5.00
Begin VB.Form frmNovember 
   BackColor       =   &H000080FF&
   Caption         =   "November Schedule"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNeeds 
      Caption         =   "What do I need to Practice?"
      Height          =   1215
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdNovember 
      Caption         =   "Go to December Schedule"
      Height          =   1335
      Left            =   960
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblNovember 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Practice Begins 11/20 at 7:00am"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   4515
      Left            =   4200
      Picture         =   "frmNovember.frx":0000
      Top             =   960
      Width           =   5385
   End
End
Attribute VB_Name = "frmNovember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNeeds_Click()
frmNovember.Hide
frmPractice.Show
End Sub

Private Sub cmdNovember_Click()
frmNovember.Hide
frmDecember.Show
End Sub
