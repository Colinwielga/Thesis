VERSION 5.00
Begin VB.Form frmMeettheMembers 
   Caption         =   "frmMeettheMembers"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   Picture         =   "frmMeettheMembers.frx":0000
   ScaleHeight     =   9990
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00000080&
      Caption         =   "Push To Start"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdAlphabetical 
      BackColor       =   &H00000080&
      Caption         =   "Sort Members in Alphabetical Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearchGrades 
      BackColor       =   &H00000080&
      Caption         =   "Search for a Specific Grade"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdNames 
      BackColor       =   &H00000080&
      Caption         =   "Search for a Specific Name"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdGrade 
      BackColor       =   &H00000080&
      Caption         =   "Sort Members in Order by Grade"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   10320
      ScaleHeight     =   5955
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   2280
      Width           =   4455
   End
   Begin VB.CommandButton cmdReturn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to the Main Screen"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Label lblMeet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Meet the Members"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   12735
   End
End
Attribute VB_Name = "frmMeettheMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: CSB/SJU Crew
'Form name: frmMeettheMembers
'Authors: Lauren Nephew and Rachel Stalley
'Date: October 18th, 2009
'Objective: To allow the user to sort the members by grade and alphabetically, as well as allow the user to search for a specific grade or name.
Option Explicit

Private Sub cmdAlphabetical_Click()

picResults.Cls   'Clears the pic box so the sorted list can be added
Dim pass As Integer, pos As Integer, j As Integer
Dim tempMembers As String, tempGrades As String, tempGender As String
picResults.Print "Member " & "                       " & "Grade" & "                      " & "Gender"
picResults.Print "*********************************************************************"
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If Members(pos) > Members(pos + 1) Then
            tempMembers = Members(pos)
            Members(pos) = Members(pos + 1)
            Members(pos + 1) = tempMembers
            tempGrades = Grades(pos)
            Grades(pos) = Grades(pos + 1)
            Grades(pos + 1) = tempGrades
            tempGender = Gender(pos)
            Gender(pos) = Gender(pos + 1)
            Gender(pos + 1) = tempGender
        End If
    Next pos
Next pass
    For j = 1 To CTR
        picResults.Print Members(j); Tab(25); Grades(j); Tab(50); Gender(j)
    Next j
    picResults.Print

End Sub

Private Sub cmdGrade_Click()

picResults.Cls 'Clears the pic box to add the sorted list
Dim pass As Integer, pos As Integer, j As Integer
Dim tempMembers As String, tempGrades As String, tempGender As String
picResults.Print "Member " & "                       " & "Grade" & "                      " & "Gender"
picResults.Print "*********************************************************************"
'Sorts members by grade
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If Grades(pos) < Grades(pos + 1) Then
            tempGrades = Grades(pos)
            Grades(pos) = Grades(pos + 1)
            Grades(pos + 1) = tempGrades
            tempMembers = Members(pos)
            Members(pos) = Members(pos + 1)
            Members(pos + 1) = tempMembers
            tempGender = Gender(pos)
            Gender(pos) = Gender(pos + 1)
            Gender(pos + 1) = tempGender
        End If
    Next pos
Next pass
    For j = 1 To CTR
                                        'The Print Left(Grades(j), 2) tells the program to print the first 2 letters of the grade name.
        picResults.Print Members(j); Tab(25); Left(Grades(j), 2); Tab(50); Gender(j)
    Next j
    picResults.Print

End Sub

Private Sub cmdNames_Click()

'Searches for a specific name entered into an input box
    Dim SearchNames As String
    SearchNames = InputBox("Enter a name (first and last) to search for a member on the crew team:", "Crew Member Search")
    Dim Found As Boolean, j As Integer
    j = 0
Do While Not Found And j < CTR
    j = j + 1
    If Members(j) = SearchNames Then
        Found = True
    End If
Loop

If (Not Found) Then
    MsgBox SearchNames & " is not a member of the Crew Team.", , "Sorry"
Else
     MsgBox SearchNames & " is a member of the Crew Team.", , "Results"
End If

    
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
'this brings the user back the the main menu scree
frmCSBSJUCrewMain.Show
frmMeettheMembers.Hide
End Sub



Private Sub cmdSearchGrades_Click()

    Dim Found As Boolean
    Dim j As Integer, Grade As String, GradeCtr As Integer
    Grade = InputBox("Enter a grade level (Freshman, Sophomore...) to see the number of members in that grade", , "Grade")
    GradeCtr = 0
    Found = False
    For j = 1 To CTR            'Searches the whole list for grade level
        If Grade = Grades(j) Then
            Found = True
            GradeCtr = GradeCtr + 1
        End If
    Next j
    If (Not Found) Then
        MsgBox Grade & " is not a grade level found on the crew team.", , "Sorry"
    Else
        MsgBox "There are " & GradeCtr & " crew team members from the " & Grade & " class.", , "Results"
    End If

End Sub


Private Sub cmdStart_Click()
cmdSearchGrades.Enabled = False
cmdStart.Enabled = True
cmdNames.Enabled = False
cmdAlphabetical.Enabled = False
cmdGrade.Enabled = False
'this button reads the file into arrays
Open App.Path & "\Members.txt" For Input As #1 'This opens the file

CTR = 0 'set the counter to 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Members(CTR), Grades(CTR), Gender(CTR)
'picResults.Print Members(CTR)
Loop
Close #1 'This closes the file
cmdSearchGrades.Enabled = True
cmdStart.Enabled = False
cmdNames.Enabled = True
cmdAlphabetical.Enabled = True
cmdGrade.Enabled = True
End Sub


