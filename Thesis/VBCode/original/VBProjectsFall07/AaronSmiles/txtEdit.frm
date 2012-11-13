VERSION 5.00
Begin VB.Form frmTxtEdit 
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmtxtEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Dim IntFile As Integer, NoteFile As String, Ctr As Integer

IntFile = FreeFile
Open App.Path & "\Notefile.txt" For Append As IntFile
    If EOF(IntFile) = False Then
        Print #IntFile, ""; Text1.Text; Date; Time
        Close #IntFile
    Else
Close IntFile
    Open App.Path & "\NoteFile.txt" For Output As IntFile
        Print #IntFile, ""; IntFile; Text1.Text; Date, Time
    Close IntFile
End If
frmtxtEdit.Hide
End Sub

Private Sub Text1_Change()

End Sub
