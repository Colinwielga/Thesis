VERSION 5.00
Begin VB.Form frmInfoKennedy 
   BackColor       =   &H00008000&
   Caption         =   "Bob Kennedy"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label lblKennedy 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoKennedy.frx":0000
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
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   240
      Picture         =   "frmInfoKennedy.frx":0370
      Top             =   240
      Width           =   3285
   End
End
Attribute VB_Name = "frmInfoKennedy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoKennedy.Hide
End Sub
