VERSION 5.00
Begin VB.Form frmInfoZatopek 
   BackColor       =   &H00008000&
   Caption         =   "Emil Zatopek"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblZatopek 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoZatopek.frx":0000
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
      Height          =   6855
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   4140
      Left            =   360
      Picture         =   "frmInfoZatopek.frx":0364
      Top             =   600
      Width           =   2580
   End
End
Attribute VB_Name = "frmInfoZatopek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoZatopek.Hide
End Sub
