VERSION 5.00
Begin VB.Form FirstForm 
   BackColor       =   &H0000FFFF&
   Caption         =   "The ACT Examiner by Derick Dehmer"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   FillColor       =   &H00FFFF80&
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton compairbox 
      Caption         =   "Look at celebrites"
      Height          =   1335
      Left            =   360
      TabIndex        =   6
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton satbutton 
      BackColor       =   &H00FF00FF&
      Caption         =   "Compute your predicted SAT score"
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton clearbox 
      BackColor       =   &H0000FF00&
      Caption         =   "Clear"
      Height          =   2415
      Left            =   6120
      MaskColor       =   &H0000FF00&
      TabIndex        =   4
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton percentbox 
      Caption         =   "Find your percent tile"
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton celbsbox 
      Caption         =   "Click here to compare with celebrities"
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Actbox 
      BackColor       =   &H0000FF00&
      Caption         =   "Click here to input name and ACT score"
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.PictureBox results 
      BackColor       =   &H00FFFF00&
      Height          =   3855
      Left            =   2280
      ScaleHeight     =   3795
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "FirstForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The ACT Examiner "M:\CS130\Project\TheACTExaminer.vbp"
'The ACT Examiner "M:\CS130\Project\TheACTExaminerI.frm"
'by Derick Dehmer
'Writen on 10/29/03
'The purpose of the ACT Examiner is to solve the
'problems studends face in interpreting the ACT
'exam scores

Option Explicit
Dim Path As String
Dim SATscore As Integer ' predicted SAT score
Dim username As String
Dim n As Double
Dim tempb As Double
Dim tempa As String
Dim pass As Double
Dim i As Double
Dim Famouspeople(1 To 12) As String ' name of celebrity
Dim famousactscores(1 To 12) As Integer ' ACT score of celebrity
Dim actscore As Integer


Private Sub Actbox_Click()
'Receiving input on user Name and ACT score
username = InputBox("Please enter your name", , "Name Please")
actscore = InputBox("Please enter your ACT score", , "ACT score")
'Test a limit on the user ACT score (range 1 to 36)
If actscore > 36 Then
    MsgBox "That is an invalid ACT score", , "Error"
End If

If actscore < 0 Then
    MsgBox "That is an invalid ACT score", , "Error"
End If
End Sub

Private Sub celbsbox_Click()
'read famous person act score from file
Open "M:\CS130\Labs\conversionfactors\celbsactscores.txt" For Input As #1
For i = 1 To 11
    Input #1, Famouspeople(i), famousactscores(i)
Next i
'Input user's name and score into the array
Famouspeople(12) = username
famousactscores(12) = actscore
n = 12
'sort names of celebrities in alphabetical order
For pass = 1 To n
    For i = 1 To n - pass
        If famousactscores(i) < famousactscores(i + 1) Then
            tempb = famousactscores(i)
            famousactscores(i) = famousactscores(i + 1)
            famousactscores(i + 1) = tempb
            tempa = Famouspeople(i)
            Famouspeople(i) = Famouspeople(i + 1)
            Famouspeople(i + 1) = tempa
        End If
    Next i
Next pass
        

results.Print Tab(5); "Celebrities"; Tab(40); "Act Score"
results.Print "--------------------------------------------------------------------------------------------------------------"
For i = 1 To 12
    results.Print Tab(5); Famouspeople(i), Tab(40); famousactscores(i)
Next i
End Sub

Private Sub clearbox_Click()
results.Cls
End Sub

Private Sub compairbox_Click()
FirstForm.Hide
secondform.Show
End Sub

Private Sub percentbox_Click()
'Tell user their percent tile based on ACT results
If actscore > 35 Then
     MsgBox "you are in the 99.999 percent tile", , "Percent tile"
ElseIf actscore > 34 Then
    MsgBox " You are in the 99.9 percent tile", , "Percent tile"
ElseIf actscore > 33 Then
    MsgBox " You are in the 99 percent tile", , "Percent tile"
ElseIf actscore > 32 Then
    MsgBox " You are in the 98.04 percent tile", , "Percent tile"
ElseIf actscore > 31 Then
    MsgBox " You are in the 96.98 percent tile", , "Percent tile"
ElseIf actscore > 30 Then
    MsgBox " You are in the 93.13 percent tile", , "Percent tile"
ElseIf actscore > 29 Then
    MsgBox " You are in the 81.90 percent tile", , "Percent tile"
ElseIf actscore > 28 Then
    MsgBox " You are in the 87.72 percent tile", , "Percent tile"
ElseIf actscore > 27 Then
    MsgBox " You are in the 80.08 percent tile", , "Percent tile"
ElseIf actscore > 26 Then
    MsgBox " You are in the 76.12 percent tile", , "Percent tile"
ElseIf actscore > 25 Then
    MsgBox " You are in the 71.96 percent tile", , "Percent tile"
ElseIf actscore > 24 Then
    MsgBox " You are in the 65.21 percent tile", , "Percent tile"
ElseIf actscore > 23 Then
    MsgBox " You are in the 58.09 percent tile", , "Percent tile"
ElseIf actscore > 22 Then
    MsgBox " You are in the 51.80 percent tile", , "Percent tile"
ElseIf actscore > 21 Then
    MsgBox " You are in the 42.67 percent tile", , "Percent tile"
ElseIf actscore > 20 Then
     MsgBox " You are in the 37.34 percent tile", , "Percent tile"
ElseIf actscore > 19 Then
    MsgBox " You are in the 31.69 percent tile", , "Percent tile"
ElseIf actscore > 18 Then
    MsgBox " You are in the 26.81 percent tile", , "Percent tile"
ElseIf actscore > 17 Then
    MsgBox " You are in the 20.27 percent tile", , "Percent tile"
ElseIf actscore > 16 Then
    MsgBox " You are in the 14.64 percent tile", , "Percent tile"
Else
     MsgBox " Your score is too low, You probably don't even know what this means", , "Percent tile"

  
End If
End Sub

Private Sub satbutton_Click()
'tell user their predicted SAT score based on ACT score
SATscore = actscore / 36 * 1600
results.Print "It is predicted that with a ACT score of "; actscore;
results.Print " you will get a"; SATscore; "on you SAT."
End Sub
