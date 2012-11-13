VERSION 5.00
Begin VB.Form frmInfoViren 
   BackColor       =   &H00008000&
   Caption         =   "Lasse Viren"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblViren 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoViren.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7095
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   3210
      Left            =   360
      Picture         =   "frmInfoViren.frx":02DA
      Top             =   720
      Width           =   2700
   End
End
Attribute VB_Name = "frmInfoViren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoViren.Hide
End Sub
