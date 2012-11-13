VERSION 5.00
Begin VB.Form frmTrivia 
   BackColor       =   &H8000000D&
   Caption         =   "So you're a trivia Buff, eh?"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Author then year here:"
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtAnswer 
      Height          =   975
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   3735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "This is far too hard for me.  I quit this trivia."
      Height          =   855
      Left            =   5400
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton CmdToAuthors 
      Caption         =   "Try out some Author Questions"
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H8000000D&
      Caption         =   "Return to Beginning"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   3120
      Width           =   1575
   End
   Begin VB.OLE OLE1 
      Class           =   "Word.Document.8"
      Height          =   1575
      Left            =   1680
      OleObjectBlob   =   "frmTrivia.frx":0000
      TabIndex        =   5
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim writers(1 To 100) As String
Dim years(1 To 100) As String
Dim works(1 To 100) As String
Dim author As String
Dim year As Integer
Dim work As String
Private Sub cmdEnter_Click()
Dim Found As Boolean, Counter As Integer, Size As Integer
author = txtAnswer.Text
Found = False
Do While Found = False And Counter < Size 'run program until found is true and counter is the same as size
    Counter = Counter + 1
    If writers(Counter) = author And years(Counter) = year Then 'when the input matches the file, display it
        Found = True
    End If
Loop
If Found = True Then
    MsgBox "You got it right! Good job!", , "Impressive"
ElseIf Found = False Then
    MsgBox "Not on top of your game. Study and try again.", , "Try again!"
End If
    
    
End Sub



Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturn_Click()
    frmWhat.Visible = True
    frmTrivia.Visible = False
    
End Sub

Private Sub CmdToAuthors_Click()
    frmAuthors.Visible = True
    frmTrivia.Visible = False
    
End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub
