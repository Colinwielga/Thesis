VERSION 5.00
Begin VB.Form frmPlayerList 
   BackColor       =   &H80000001&
   Caption         =   "Player List"
   ClientHeight    =   5895
   ClientLeft      =   4185
   ClientTop       =   3105
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   5805
   Begin VB.PictureBox picResults 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   480
      ScaleHeight     =   4455
      ScaleWidth      =   4815
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Title Screen"
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   4920
      Width           =   2295
   End
End
Attribute VB_Name = "frmPlayerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdReturn_Click()
frmTitleScreen.Show
frmPlayerList.Hide
End Sub

Private Sub Form_activate()
Dim ctr As Integer
Dim pos As Integer
Dim First(1 To 100) As String
Dim Last(1 To 100) As String

picResults.Print "This is a list of players who have recently completed the Game:"
picResults.Print "---------------------------------------------------------------"
Open App.Path & "\Completion.txt" For Input As #1

ctr = 0
Do Until EOF(1)
ctr = ctr + 1
Input #1, First(ctr), Last(ctr)
Loop
Close #1

For pos = 1 To ctr
    picResults.Print First(pos) & " " & Last(pos)
Next pos

End Sub
