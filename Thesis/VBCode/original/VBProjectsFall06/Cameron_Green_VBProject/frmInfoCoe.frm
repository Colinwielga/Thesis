VERSION 5.00
Begin VB.Form frmInfoCoe 
   BackColor       =   &H00008000&
   Caption         =   "Sebastian Coe"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label lblCoe 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoCoe.frx":0000
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
      Height          =   7335
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   4245
      Left            =   600
      Picture         =   "frmInfoCoe.frx":037D
      Top             =   1080
      Width           =   2985
   End
End
Attribute VB_Name = "frmInfoCoe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoCoe.Hide
End Sub
