VERSION 5.00
Begin VB.Form frmInfoBekele 
   BackColor       =   &H00008000&
   Caption         =   "Kenenisa Bekele"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label lblBekele 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoBekele.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   480
      Picture         =   "frmInfoBekele.frx":025F
      Top             =   1080
      Width           =   2280
   End
End
Attribute VB_Name = "frmInfoBekele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoBekele.Hide
End Sub

