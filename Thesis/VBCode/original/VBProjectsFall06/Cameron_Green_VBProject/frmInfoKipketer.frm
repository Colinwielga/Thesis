VERSION 5.00
Begin VB.Form frmInfoKipketer 
   BackColor       =   &H00008000&
   Caption         =   "Wilson Kipketer"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label lblKipketer 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoKipketer.frx":0000
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
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   3405
      Left            =   480
      Picture         =   "frmInfoKipketer.frx":0201
      Top             =   1080
      Width           =   2250
   End
End
Attribute VB_Name = "frmInfoKipketer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoKipketer.Hide
End Sub
