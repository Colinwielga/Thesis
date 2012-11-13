VERSION 5.00
Begin VB.Form frmRange 
   BackColor       =   &H00400000&
   Caption         =   "Rangeof Scores"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to main page"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdRange 
      Caption         =   "Range "
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picRange 
      BackColor       =   &H00C0E0FF&
      Height          =   5055
      Left            =   1800
      ScaleHeight     =   4995
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this page allows the user to load all of Shaun's competition scores.
'It then arranges them from smallest to largest so the user can see the complete
'range of scores Shaun has received in competition

'this function loads the data from the text file into an array. It then compares the
'first value with the second value for both arrays. If the first value>second value
'it then swaps the value places. It continues this until the end of the file is reached
'and the data is sorted from smallest to largest, at which point the results are printed.
Private Sub cmdRange_Click()
    Open App.Path & "\compiled.txt" For Input As #1
    ctr = 0
    
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, Ran1(ctr), Ran2(ctr)
    Loop
    Close #1
    
    For Pass = 1 To (ctr - 1)
        For Pos = 1 To ctr - Pass
            If Ran1(Pos) > Ran1(Pos + 1) Then
                TempRun1 = Ran1(Pos)
                Ran1(Pos) = Ran1(Pos + 1)
                Ran1(Pos + 1) = TempRun1
            End If
        Next Pos
    Next Pass
    
    For Pass = 1 To (ctr - 1)
        For Pos = 1 To ctr - Pass
            If Ran2(Pos) > Ran2(Pos + 1) Then
                TempRun2 = Ran2(Pos)
                Ran2(Pos) = Ran2(Pos + 1)
                Ran2(Pos + 1) = TempRun2
            End If
        Next Pos
    Next Pass
    picRange.Cls
    picRange.Print "Shaun's range of scores for all competition is:"
    picRange.Print "                                      "
    picRange.Print "Run 1", "Run 2"
    For Pos = 1 To (ctr)
        picRange.Print Ran1(Pos), Ran2(Pos)
    Next Pos
        
End Sub
'returns user to main page
Private Sub cmdReturn_Click()
    frmShaunWhite.Show
    frmRange.Hide
End Sub
