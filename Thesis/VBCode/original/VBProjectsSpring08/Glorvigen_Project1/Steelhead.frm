VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00800080&
   Caption         =   "Form10"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form10"
   ScaleHeight     =   4965
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00800080&
      Caption         =   "Back To Main Page"
      Height          =   855
      Left            =   5760
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "Steelhead"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
    Form1.Show
    Form10.Hide
    
End Sub

Private Sub Form_Load()

    If age < 16 Then
        MsgBox "You Will Not need to buy a Trout Stamp", , "Attention"
    Else
        MsgBox ("You Will Need to purchase a Trout Stamp Before Trouting Fishing")
    End If
    
        
    
End Sub
