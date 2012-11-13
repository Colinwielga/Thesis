VERSION 5.00
Begin VB.Form frmInfoMax 
   BackColor       =   &H00008000&
   Caption         =   "What is VO2 Max?"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to VO2 Max Page"
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblMax 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoMax.frx":0000
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   7455
   End
End
Attribute VB_Name = "frmInfoMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to VO2 max page, information on form tells what VO2 max is'
Private Sub cmdBack_Click()
    frmInfoMax.Hide
    frmVO2Max.Show
End Sub
