VERSION 5.00
Begin VB.Form MeetTheTeamForm 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   9540
   ClientLeft      =   1650
   ClientTop       =   1245
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MousePointer    =   12  'No Drop
   ScaleHeight     =   9540
   ScaleWidth      =   12210
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Team Info"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort by name"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H008080FF&
      Caption         =   "Search for someone"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H000080FF&
      Caption         =   "Read the Team"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      MaskColor       =   &H000000FF&
      MousePointer    =   13  'Arrow and Hourglass
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FFFF&
      Height          =   4215
      Left            =   2040
      ScaleHeight     =   4155
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Image teampic 
      Height          =   4500
      Left            =   960
      Picture         =   "MeetTheTeam.frx":0000
      Top             =   4680
      Width           =   6750
   End
End
Attribute VB_Name = "MeetTheTeamForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WaterPoloProject
'MeetTheTeamForm
'Bobby Chapman
'Written 3/14/2009
'Objective- to read 3 arrarys and print the total number of goals and ejections, and
'search the array for a team member, and to put the team in alphabetical order

Option Explicit
'declare global variables
Dim Names(1 To 15) As String, Goals(1 To 15) As Integer
Dim Ejections(1 To 15) As Integer, Ctr As Integer

Private Sub cmdRead_Click()
'declare local variables
Dim TotalGoals As Integer, TotalEjections As Integer

'sets ctrs to 0
Ctr = 0
TotalGoals = 0
TotalEjections = 0

'this opens the document
Open App.Path & "\TheTeam.txt" For Input As #1

'prints the header info
picResults.Print Tab(5); "Name"; Tab(30); "Goals"; Tab(40); "Ejections"
picResults.Print "********************************************************************"


Do While Not EOF(1)
    'increases ctr each time through the loops
    Ctr = Ctr + 1
    'reads the data into 3 arrarys
    Input #1, Names(Ctr), Goals(Ctr), Ejections(Ctr)
    'prints the arrays
    picResults.Print Names(Ctr); Tab(30); Goals(Ctr); Tab(40); Ejections(Ctr)
    'calculates total goals
    TotalGoals = TotalGoals + Goals(Ctr)
    'calculates total ejections
    TotalEjections = TotalEjections + Ejections(Ctr)
Loop

'prints the total goals and ejections
picResults.Print "Total"; Tab(30); TotalGoals; Tab(40); TotalEjections

'closes the file used for input
Close #1

'makes the Search and Sort buttons visible
cmdSearch.Visible = True
cmdSort.Visible = True
cmdRead.Visible = False

End Sub

Private Sub cmdSearch_Click()
'declare local variables
Dim Search As String, Found As Boolean, I As Integer

'sets found to false. to be used when searching
Found = False
'sets I to 0 to be used in search
I = 0

'enter a name from an input box to search if on team
Search = InputBox("Enter a name to find out if they're on the team", "Search")

'searches the array to see if the searched person is on the team
Do While Not Found And I <= Ctr
I = I + 1
    If Search = Names(I) Then
        Found = True
    End If
Loop

If Search = "" Then
    Found = False
    End If

'if the person is found, it shows a message box saying they're on the team
'if not found, displays message box saying that they aren't found
If Found Then
    MsgBox "Sweet, " & Search & " is on the team", , "Woot!"
        Else
            MsgBox "Sorry, but " & Search & " isn't on the team", , "Alert"
End If

End Sub

Private Sub cmdSort_Click()
'declare local variables
Dim Pass As Integer, Pos As Integer, TempNames As String, I As Integer
Dim TempGoals As Integer, TempEjections As Integer

'clears the picResults box
picResults.Cls

'prints the header
picResults.Print "Name"; Tab(30); "Goals"; Tab(40); "Ejections"
picResults.Print "********************************************************"

'sorts the the arrays into alphabetical order
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Names(Pos) > Names(Pos + 1) Then
            TempNames = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = TempNames
            
            TempGoals = Goals(Pos)
            Goals(Pos) = Goals(Pos + 1)
            Goals(Pos + 1) = TempGoals
            
            TempEjections = Ejections(Pos)
            Ejections(Pos) = Ejections(Pos + 1)
            Ejections(Pos + 1) = TempEjections
        End If
    Next Pos
Next Pass

'prints the arrays in alphabetical order
For I = 1 To Ctr
    picResults.Print Names(I); Tab(30); Goals(I); Tab(40); Ejections(I)
Next I
End Sub

Private Sub cmdReturn_Click()
'goes back to team form
MeetTheTeamForm.Hide
TeamForm.Show
End Sub
