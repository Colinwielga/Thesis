VERSION 5.00
Begin VB.Form frmRock 
   BackColor       =   &H00004040&
   Caption         =   "Rocks"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next!"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      Height          =   2895
      Left            =   1680
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Find!"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frmRock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNext_Click()
    frmNumber.Show
    frmRock.Hide

End Sub

Private Sub cmdSort_Click()
    Dim Count As Integer
    Dim Pos As Integer
    Dim Found As Boolean
    Dim Rname As String
    Dim Rocks(1 To 100) As String
    Dim Rock As String
        Open App.Path & "\rocks.txt" For Input As #1
    Count = 0
        Do Until EOF(1)
            Input #1, Rocks
            Count = Count + 1
            Rocks(Count) = Rock
        Loop
    Close #1
    Rname = InputBox("Input the name of the rock you want to find", "Rock")
    
    Found = False
    Pos = 0
    
    Do While (Found = False And Pos < Count)
        Pos = Pos + 1
        If Rocks(Pos) = Rname Then
            Found = True
        End If
    Loop
    
    If Found = True Then
        picResults.Print Rname; " is in this list of rocks"; Pos
    Else
        picResults.Print Rname; " is not in this list of rocks"
    End If
End Sub
