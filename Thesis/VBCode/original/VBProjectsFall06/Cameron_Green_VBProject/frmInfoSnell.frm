VERSION 5.00
Begin VB.Form frmInfoSnell 
   BackColor       =   &H00008000&
   Caption         =   "Peter Snell"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblSnell 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoSnell.frx":0000
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
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   3450
      Left            =   360
      Picture         =   "frmInfoSnell.frx":0328
      Top             =   1080
      Width           =   3270
   End
End
Attribute VB_Name = "frmInfoSnell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoSnell.Hide
End Sub
