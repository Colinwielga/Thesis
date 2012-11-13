VERSION 5.00
Begin VB.Form frmInfoRyun 
   BackColor       =   &H00008000&
   Caption         =   "Jim Ryun"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label lblRyun 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoRyun.frx":0000
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
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   5940
      Left            =   240
      Picture         =   "frmInfoRyun.frx":03BB
      Top             =   120
      Width           =   4275
   End
End
Attribute VB_Name = "frmInfoRyun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoRyun.Hide
End Sub
