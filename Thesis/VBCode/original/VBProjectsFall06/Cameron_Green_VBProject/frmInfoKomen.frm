VERSION 5.00
Begin VB.Form frmInfoKomen 
   BackColor       =   &H00008000&
   Caption         =   "Daniel Komen"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label lblKomen 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoKomen.frx":0000
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
      Height          =   7215
      Left            =   4440
      TabIndex        =   1
      Top             =   240
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "frmInfoKomen.frx":036D
      Top             =   120
      Width           =   4155
   End
End
Attribute VB_Name = "frmInfoKomen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoKomen.Hide
End Sub
