VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roster"
   ClientHeight    =   5040
   ClientLeft      =   5385
   ClientTop       =   3480
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5565
   Begin VB.CommandButton cmdReturnHome 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Home Page"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox picRoster 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   2160
      ScaleHeight     =   4335
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H000000FF&
      Caption         =   "Display"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisplay_Click()
'this sub routine displays the names of the team into an array
Open App.Path & "\TeamNames.txt" For Input As #1

Dim Swimmer(1 To 100) As String
Dim Ctr As Integer
picRoster.Cls
Ctr = 0

Do Until EOF(1)
Ctr = Ctr + 1
Input #1, Swimmer(Ctr)
Loop

For J = 1 To Ctr

picRoster.Print Tab(10); Swimmer(J)
Next J

Close #1

End Sub

Private Sub cmdReturnHome_Click()
frmRoster.Hide
frmHomePage.Show

End Sub
