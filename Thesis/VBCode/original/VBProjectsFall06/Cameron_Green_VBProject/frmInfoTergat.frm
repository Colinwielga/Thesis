VERSION 5.00
Begin VB.Form frmInfoTergat 
   BackColor       =   &H00008000&
   Caption         =   "Paul Tergat"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label lblTergat 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoTergat.frx":0000
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
      Height          =   6975
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   8775
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   240
      Picture         =   "frmInfoTergat.frx":0433
      Top             =   240
      Width           =   3780
   End
End
Attribute VB_Name = "frmInfoTergat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoTergat.Hide
End Sub
