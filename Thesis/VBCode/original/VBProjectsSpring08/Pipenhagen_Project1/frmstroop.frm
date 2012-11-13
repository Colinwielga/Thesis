VERSION 5.00
Begin VB.Form frmstroop 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   6
      Top             =   7560
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      Height          =   1815
      Left            =   2040
      ScaleHeight     =   1755
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   2040
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Trial"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdred 
      Caption         =   "Red"
      Height          =   855
      Left            =   1200
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdblue 
      Caption         =   "Blue"
      Height          =   855
      Left            =   3720
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdgreen 
      Caption         =   "Green"
      Height          =   855
      Left            =   6240
      TabIndex        =   1
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   6120
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Score Your Trials"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   5880
      Width           =   4935
   End
End
Attribute VB_Name = "frmstroop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmcognitivebehavioral
'Author: Calvin Pipenhagen
'Date Written: March 26, 2008
'Objective: To provide a rough approximation of the original stroop task in the hopes of replicating the results.
Option Explicit
Dim n As Integer
Dim ctr As Integer
Dim i As Integer
Dim j As Integer
Dim temp As String
Dim match As Integer
Dim words(1 To 1000) As String
Dim mismatch As Integer
Dim time As Integer
Dim matchtime As Integer
Dim mismatchtime As Integer
Private Sub cmdblue_Click() 'Determines if the color of the word matches the word itself. For instance the word "blue" may be the color green
Timer1.Enabled = False      'this would be classified as a mismatch. If the word and the color of the word the same they are considered a match.
If words(1) = "blue1" Then  'The timer ends, it has counted the length of time since the user pressed the begin trial button
    match = match + 1 'counts the number of matches
    matchtime = matchtime + time 'keeps a running total of how much time task has taken
End If
If words(1) = "red3" Then 'if blue is clicked and the word loaded was green this condition occurs
    mismatch = mismatch + 1 'counts number of mismatches
    mismatchtime = mismatchtime + time 'keeps a running total of mismatch time
End If
If words(1) = "green3" Then 'if blue was red
    mismatch = mismatch + 1
    mismatchtime = mismatchtime + time
End If
time = 0 'timer is reset before the beginning of the next trial
End Sub

Private Sub cmdgreen_Click() 'Determines the various scores for green colored words
Timer1.Enabled = False
If words(1) = "green1" Then
    match = match + 1
    matchtime = matchtime + time
End If
If words(1) = "blue2" Then
    mismatch = mismatch + 1
    mismatchtime = mismatchtime + time
End If
If words(1) = "red2" Then
    mismatch = mismatch + 1
    mismatchtime = mismatchtime + time
End If
time = 0
End Sub

Private Sub cmdred_Click() 'determines the various scores for red colored words
Timer1.Enabled = False
If words(1) = "red1" Then
    match = match + 1
    matchtime = matchtime + time
End If
If words(1) = "blue3" Then
    mismatch = mismatch + 1
    mismatchtime = mismatchtime + time
End If
If words(1) = "green2" Then
    mismatch = mismatch + 1
    mismatchtime = mismatchtime + time
End If
time = 0
End Sub


Private Sub Command1_Click()
Timer1.Enabled = True 'the timer starts at zero when this button is pressed
time = 0
n = n + 1 'The nine variations of words that are loaded at a form level are here chosen randomly
    For i = 1 To 9
        j = Int((9 - i + 1) * Rnd + i)
        temp = words(i)
        words(i) = words(j)
        words(j) = temp
    Next i
 picresults.Picture = LoadPicture(App.Path & "\stroop\" & words(1) & ".jpg") 'the colored word is loaded from a file
End Sub

Private Sub cmdscore_click()
If (matchtime / match) / 10 < (mismatchtime / mismatch) / 10 Then 'determines if the results were replicated for this to work both a match and a mismatch must occur
    MsgBox "You succesfully replicated the results of the original Stroop task. When the ink color and color word matched your responses were significantly faster than when there was a discrepency." & FormatNumber((matchtime / match) / 10, 2) & "seconds compared to" & FormatNumber((mismatchtime / mismatch) / 10, 2) & "seconds.", , "Results"
Else
    MsgBox "Your results failed to replicate those of the original Stroop task."
End If
mismatchtime = 0 'resets the various counters so the experiment can be started again
matchtime = 0
match = 0
mismatch = 0
End Sub

Private Sub Command2_Click() 'goes back to main cognitive behavioral page
frmcognitivebehavioral.Show
frmstroop.Hide
End Sub

Private Sub Form_Load()
MsgBox "The stroop task is a classic cognitive psychological experiment. Stroop (1935) showed participants color words sometimes printed in an alternate color. For instance, the color word blue could be printed in red ink. Participants were instructed to state the color of the word as quickly as possible. When words the color word and the ink matched, this task was relatively easy. However, when there was a discrepancy participants struggled to name the color of the word. These results provided insight into automatic processing. Ultimately, it is difficult to not read the color word. For the following demo you are asked to click on the color that corresponds to the ink color of a presented word. To begin a trial simply click the begin trial button.", , "Information and Instructions"

Open App.Path & "\randomizer.txt" For Input As #1 'this loads a series of words from a data file. These words are later used as part of a file name to load a corresponding picture
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, words(ctr)
Loop
Close #1
End Sub
Private Sub Timer1_Timer() 'A timer is establised. This will be used to determine the amount of time it takes for the user to respond
time = time + 1 'the timer increases by one
End Sub

