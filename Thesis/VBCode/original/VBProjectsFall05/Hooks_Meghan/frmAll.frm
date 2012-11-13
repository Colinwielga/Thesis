VERSION 5.00
Begin VB.Form frmAll 
   BackColor       =   &H00FF8080&
   Caption         =   "All Facts"
   ClientHeight    =   3675
   ClientLeft      =   2580
   ClientTop       =   2790
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   9960
   Begin VB.CommandButton cmdOK2 
      Caption         =   "OK"
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "View All"
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.PictureBox picAll 
      BackColor       =   &H00FFC0FF&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All (frmAll)
'Created by Meghan Hooks
'10-30-05
    'The purpose of this form is to show all of the facts (in alphabetical order)
    'so the user doesn't have to click throught them randomly as with the
    'Fact of the Day' button

Private Sub Picture1_Click()

End Sub

Private Sub cmdOK2_Click()
frmAll.Hide
frmProject.Show
End Sub

Private Sub cmdShow_Click()
Dim K As Integer
Dim Pass As Integer
Dim Q As Integer
Dim View(1 To 9) As String
Dim D As Integer
Dim Temp As String

Open App.Path & "\TIPOFDAY.txt" For Input As #1
For K = 1 To 9
Input #1, View(K)
Next K
Close #1

For Pass = 1 To 8
    For Q = 1 To 9 - Pass
        If View(Q) > View(Q + 1) Then
            Temp = View(Q)
            View(Q) = View(Q + 1)
            View(Q + 1) = Temp
        End If
    Next Q
Next Pass

For W = 1 To 9
picAll.Print View(W)
Next W

End Sub

