VERSION 5.00
Begin VB.Form frmmajrlist 
   BackColor       =   &H00000080&
   Caption         =   "Form2"
   ClientHeight    =   8895
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   Picture         =   "majrlist.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   13440
      TabIndex        =   31
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdtheology 
      Caption         =   "Theology"
      Height          =   615
      Left            =   4440
      TabIndex        =   30
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton cmdtheater 
      Caption         =   "Theater"
      Height          =   615
      Left            =   4440
      TabIndex        =   29
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdsociology 
      Caption         =   "Sociology"
      Height          =   615
      Left            =   4440
      TabIndex        =   28
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdpsychology 
      Caption         =   "Psychology"
      Height          =   615
      Left            =   4440
      TabIndex        =   27
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdpoliticalscience 
      Caption         =   "Political Science"
      Height          =   615
      Left            =   4440
      TabIndex        =   26
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdphysics 
      Caption         =   "Physics"
      Height          =   735
      Left            =   4440
      TabIndex        =   25
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdphilosophy 
      Caption         =   "Philosophy"
      Height          =   735
      Left            =   4440
      TabIndex        =   24
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdpeacestudies 
      Caption         =   "Peace Studies"
      Height          =   735
      Left            =   4440
      TabIndex        =   23
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdnutrition 
      Caption         =   "Nutrition"
      Height          =   735
      Left            =   4440
      TabIndex        =   22
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdnursing 
      Caption         =   "Nursing"
      Height          =   735
      Left            =   4440
      TabIndex        =   21
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdmusic 
      Caption         =   "Music"
      Height          =   735
      Left            =   2400
      TabIndex        =   20
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdmathematics 
      Caption         =   "Mathematics"
      Height          =   735
      Left            =   2400
      TabIndex        =   19
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdmanagement 
      Caption         =   "Management"
      Height          =   615
      Left            =   2400
      TabIndex        =   18
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdhistory 
      Caption         =   "History"
      Height          =   615
      Left            =   2400
      TabIndex        =   17
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdhispanicstudies 
      Caption         =   "Hispanic Studies"
      Height          =   735
      Left            =   2400
      TabIndex        =   16
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdgerman 
      Caption         =   "German"
      Height          =   735
      Left            =   2400
      TabIndex        =   15
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to Home"
      Height          =   855
      Left            =   13320
      TabIndex        =   14
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrench 
      Caption         =   "French"
      Height          =   735
      Left            =   2400
      TabIndex        =   13
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdenvironmentalstudies 
      Caption         =   "Environmental Studies"
      Height          =   735
      Left            =   2400
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdenglish 
      Caption         =   "English"
      Height          =   735
      Left            =   2400
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdeducation 
      Caption         =   "Education"
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdeconomics 
      Caption         =   "Economics"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmdcomputerschience 
      Caption         =   "Computer Science"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdcommunication 
      Caption         =   "Communication"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdchemistory 
      Caption         =   "Chemistory"
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdbiology 
      Caption         =   "Biology"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdresister 
      Caption         =   "Register"
      Height          =   1335
      Left            =   13320
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox picresults 
      Height          =   8175
      Left            =   6960
      ScaleHeight     =   8115
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   600
      Width           =   6255
   End
   Begin VB.CommandButton cmdchemistory 
      Caption         =   "Biochemistory"
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmdart 
      Caption         =   "Art"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdaccount 
      Caption         =   "Accounting "
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmmajrlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdaccount_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Accounting
    Open App.Path & "\accounting major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Accounting Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"

'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdart_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Art
    Open App.Path & "\art major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Art Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdback_Click()
frmhome.Show
frmmajrlist.Hide
End Sub

Private Sub cmdbiology_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Biology
    Open App.Path & "\biology major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Biology Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdchemistory_Click(Index As Integer)
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Chemistory
    Open App.Path & "\chemistory major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Chemistory Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdclear_Click()
picresults.Cls
End Sub

Private Sub cmdcommunication_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Communication
    Open App.Path & "\communication major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Communication Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdcomputerschience_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Computer Science
    Open App.Path & "\computer science major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all AComputer Science Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdeconomics_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Economics
    Open App.Path & "\economics major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Economics Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdeducation_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Education
    Open App.Path & "\education major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Education Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdenglish_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for English
    Open App.Path & "\english major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all English Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdenvironmentalstudies_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Environmental Studies
    Open App.Path & "\environmental studies major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Environmental Studies Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdfrench_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for French
    Open App.Path & "\french major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all French Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdgender_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Gender & Women's Studies
    Open App.Path & "\gender & women's studies major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Gender Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdgerman_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for German
    Open App.Path & "\german major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all German Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdhispanicstudies_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Hispanic Studies
    Open App.Path & "\hispanic studies major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Hispanic Studies Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdhistory_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for History
    Open App.Path & "\history major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all History Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdmanagement_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Management
    Open App.Path & "\management major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Management Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdmathematics_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Mathematics
    Open App.Path & "\mathematics.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Mathematics Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdmusic_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Music
    Open App.Path & "\music major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Music Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdnursing_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Nursing
    Open App.Path & "\nursing major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Nursing Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdnutrition_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Nutrition
    Open App.Path & "\nutrition major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Nutrition Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdpeacestudies_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Peace Study
    Open App.Path & "\peace study major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Peace Studeis Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdphilosophy_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Philosophy
    Open App.Path & "\philosophy major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Philosophy Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdphysics_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Physics
    Open App.Path & "\physics major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Physics Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdpoliticalscience_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Political Science
    Open App.Path & "\political science major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Political Science Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdpsychology_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Psychology
    Open App.Path & "\psychology major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Psychology Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Option Explicit
' Written by John O'Grady and Yuzu Takei
' Written 10-17-09

Private Sub cmdresister_Click()
frmregister.Show
End Sub

Private Sub cmdsociology_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Sociology
    Open App.Path & "\sociology major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Sociology Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdtheater_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Theater
    Open App.Path & "\theater major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Theater Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

Private Sub cmdtheology_Click()
'initialize CTR to zero, to be used for the position in the array
    CTR = 0

'open the data file for Theology
    Open App.Path & "\theology major.txt" For Input As #1

'print heading for the table
    picresults.Print "Required coursework for all Theology Majors"
    picresults.Print "******************************************************"
    picresults.Print "Major", "Course No.", "Credits", "Class"
'read the numbers from the file
    Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Majors(CTR), CN(CTR), Classes(CTR), Credits(CTR)
    picresults.Print Majors(CTR), CN(CTR), Credits(CTR), Classes(CTR)
Loop
Close #1
End Sub

