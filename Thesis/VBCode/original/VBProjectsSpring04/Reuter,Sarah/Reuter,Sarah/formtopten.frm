VERSION 5.00
Begin VB.Form formtopten 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form2"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14820
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   14820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save new top ten list and exit quiz!"
      Enabled         =   0   'False
      Height          =   2415
      Left            =   7320
      Picture         =   "formtopten.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Click here to add your name to the Top Ten List!"
      Enabled         =   0   'False
      Height          =   2295
      Left            =   4320
      Picture         =   "formtopten.frx":2A0D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00C0FFFF&
      Height          =   4455
      Left            =   720
      ScaleHeight     =   4395
      ScaleWidth      =   9315
      TabIndex        =   1
      Top             =   3360
      Width           =   9375
   End
   Begin VB.CommandButton cmdclick 
      Caption         =   "Click here to see the Top Ten list!"
      Height          =   2295
      Left            =   840
      Picture         =   "formtopten.frx":56C8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Created by Sarah Reuter"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   8040
      Width           =   1815
   End
End
Attribute VB_Name = "formtopten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Sarah's Kermit the Frog Quiz
'Form name and file: formtopten(formtopten.frm)
'Created by Sarah Reuter
'Written 3/14/04
'Purpose: top ten list

Dim Names(1 To 11) As String
Dim Scores(1 To 11) As Single
Dim CTR As Integer
Dim Path As String


'get the player's name and score and adjust top ten list
Private Sub cmdadd_Click()
Newname = InputBox("Please Enter Your Name")
Newscore = Correct
If Newscore > Scores(10) Then
    Scores(10) = Newscore
    Names(10) = Newname
    picresults.Cls
    N = 10
    For Pass = 1 To N - 1
        For CTR = 1 To N - Pass
            If Scores(CTR) < Scores(CTR + 1) Then
                TempName = Names(CTR)
                Names(CTR) = Names(CTR + 1)
                Names(CTR + 1) = TempName
                TempScore = Scores(CTR)
                Scores(CTR) = Scores(CTR + 1)
                Scores(CTR + 1) = TempScore
            End If
        Next CTR
Next Pass

picresults.Print "Name"; Tab(25); "Score"
picresults.Print "*************************************************"
For CTR = 1 To 10
    picresults.Print Names(CTR); Tab(25); Scores(CTR)
Next CTR
End If
If Newscore < Scores(10) Then
    picresults.Print
    picresults.Print "I'm sorry, you did not score high enough to make it into the top ten."
End If
cmdsave.Enabled = True
End Sub

'display and sort top ten list
Private Sub cmdclick_Click()
Dim Pass As Integer
Dim TempName As String
Dim TempScore As Single
Path = "N:\CS130\handin\Reuter, Sarah\Reuter,Sarah\"
Open Path & "topten.txt" For Input As #1
For CTR = 1 To 10
    Input #1, Names(CTR), Scores(CTR)
Next CTR
Close #1
N = 10
For Pass = 1 To N - 1
    For CTR = 1 To N - Pass
        If Scores(CTR) < Scores(CTR + 1) Then
            TempName = Names(CTR)
            Names(CTR) = Names(CTR + 1)
            Names(CTR + 1) = TempName
            TempScore = Scores(CTR)
            Scores(CTR) = Scores(CTR + 1)
            Scores(CTR + 1) = TempScore
        End If
    Next CTR
Next Pass

picresults.Print "Name"; Tab(25); "Score"
picresults.Print "*************************************"
For CTR = 1 To 10
    picresults.Print Names(CTR); Tab(25); Scores(CTR)
Next CTR
cmdadd.Enabled = True
cmdclick.Enabled = False
End Sub

Private Sub cmdsave_Click()
Open Path & "topten.txt" For Output As #1
For CTR = 1 To 10
    Write #1, Names(CTR), Scores(CTR)
Next CTR
Close #1

End
End Sub
