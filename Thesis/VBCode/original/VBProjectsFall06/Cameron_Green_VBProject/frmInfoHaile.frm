VERSION 5.00
Begin VB.Form frmInfoHaile 
   BackColor       =   &H00008000&
   Caption         =   "Haile Gebrselassie"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label lblHaile 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoHaile.frx":0000
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
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   120
      Picture         =   "frmInfoHaile.frx":03FB
      Top             =   120
      Width           =   4470
   End
End
Attribute VB_Name = "frmInfoHaile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmInfoHaile.Hide
    frmRunners.Show
End Sub
