VERSION 5.00
Begin VB.Form frmEau_Claire_Results 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   Picture         =   "eau_claire_invite.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackToResults 
      Caption         =   "Back to Main Results Page"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdFullResults 
      Caption         =   "View Team Results/Top Individuals"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdTeam 
      Caption         =   "A Team"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdPlace 
      Caption         =   "A Place"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdIndividual 
      Caption         =   "An Individual"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   6120
      ScaleHeight     =   7635
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Search Results for . . ."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   6120
      Width           =   6135
   End
   Begin VB.Label lblColfax 
      BackColor       =   &H000000FF&
      Caption         =   "Colfax, WI 9/21/2007"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "UW-Eau Claire Invite"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmEau_Claire_Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CTRteam As Integer
Dim TeamPlace(1 To 1000) As Integer
Dim TeamScore(1 To 1000) As Integer
Dim TeamName(1 To 1000) As String

Dim CTR As Integer
Dim place(1 To 1000) As Integer
Dim runner(1 To 1000) As String
Dim team(1 To 1000) As String
Dim minutes(1 To 1000) As Integer
Dim seconds(1 To 1000) As String
'seconds must be declared as a string, instead of an integer.
'when declared as an integer the value 09 will be printed as 9.
'this would make a time of 26:09 appear as a time of 26:9 which is impossible.


'pressing this button takes you back to the main results page
Private Sub cmdBackToResults_Click()
    frmEau_Claire_Results.Hide
    frmSeason_Results.Show
End Sub

'Pressing the button that says "View Full Results" will load the results file
'into five array's and then print the race results in a picture box
Private Sub cmdFullResults_Click()
    Dim J As Integer
    Dim I As Integer
    
    Open App.Path & "\bluegold_invite_team_results.txt" For Input As #1
    Do Until EOF(1)
        CTRteam = CTRteam + 1
        Input #1, TeamPlace(CTRteam), TeamName(CTRteam), TeamScore(CTRteam)
    Loop
    Close (1)
    
    picResults.Print "Place"; Tab(10); "Team"; Tab(28); "Score"
    
    Do While J < CTRteam
        J = J + 1
        picResults.Print TeamPlace(J); Tab(10); TeamName(J); Tab(28); TeamScore(J)
    Loop
    
    Close (1)
    
    picResults.Print "____________________________________________________"
    
    Open App.Path & "\bluegold_invite_individual_results.txt" For Input As #2
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, place(CTR), runner(CTR), team(CTR), minutes(CTR), seconds(CTR)
    Loop
    
    picResults.Print "Place"; Tab(10); "Name"; Tab(30); "Team"; Tab(50); "Time"
    
    
    'because the picture box is not large enough, only the team results and a few of the top individuals can be printed.
    'the full results are still stored in the array, and the user can use the search buttons to find results that did not fit on the main screen
    Do While I < 15
        I = I + 1
        picResults.Print place(I); Tab(10); runner(I); Tab(33); team(I); Tab(53); minutes(I); ":"; seconds(I)
    Loop
    
    Close (2)
    
    'the button to view the full results is disabled
    'and the search buttons are now enabled to search for results that did not fit on the main screen
    cmdFullResults.Enabled = False
    cmdIndividual.Enabled = True
    cmdPlace.Enabled = True
    cmdTeam.Enabled = True
    cmdBackToResults.Enabled = True
    
End Sub

'Pressing this button allows the user to search for a certain individual.
'The individual will then appear in the box picResults, along with that person place, school, and time.
'If the user searches for a runner who did not run this race a Message Box will appear to notify the user.

Private Sub cmdIndividual_Click()
    Dim SearchRunner As String
    Dim found As Boolean
    Dim I As Integer
    Dim output As String
    
    picResults.Cls
    
    SearchRunner = InputBox("Who are you looking for?", "Enter a Name")
    
    picResults.Print "Place"; Tab(10); "Name"; Tab(33); "Team"; Tab(53); "Time"
    
    I = 0
    Do While I < CTR And Not found = True
        I = I + 1
        If SearchRunner = runner(I) Then
            
            picResults.Print place(I); Tab(10); runner(I); Tab(33); team(I); Tab(53); minutes(I); ":"; seconds(I)
            found = True
        End If
    Loop
    
    If (Not found) Then
        output = MsgBox("The Runner you searched for was not in this race")
    End If
    
End Sub

'This button allows the user to search for a certain place.
'It shows who finished in that place, along with their school, and their time.
'If the user enters a number that is greater than the number of runners, or if they enter a value that is not an integer, a message box appears to notify the user.

Private Sub cmdPlace_Click()

    Dim SearchPlace As Integer
    Dim found As Boolean
    Dim I As Integer
    Dim output As String
    
    picResults.Cls
    
    SearchPlace = InputBox("Which Finisher are you Looking For?", "Enter a Place")
    
    picResults.Print "Place"; Tab(10); "Name"; Tab(33); "Team"; Tab(53); "Time"
    
    I = 0
    Do While I < CTR And Not found = True
        I = I + 1
        If SearchPlace = place(I) Then
            picResults.Print place(I); Tab(10); runner(I); Tab(33); team(I); Tab(53); minutes(I); ":"; seconds(I)
            found = True
        End If
    Loop
    
    If (Not found) Then
        output = MsgBox("There was no finisher in that place")
    End If
    
End Sub


'This button allows the user to search for runners from a specific team.
'All runners from that team will be shown in the box picResults.
'If the user searches for a team that was not in the race, a Message Box will appear to tell the user.

Private Sub cmdTeam_Click()
    Dim SearchTeam As String
    Dim found As Boolean
    Dim I As Integer
    Dim output As String
    
    picResults.Cls
    
    SearchTeam = InputBox("Which Team are you Looking For?", "Enter a Team")
    
    picResults.Print "Place"; Tab(10); "Name"; Tab(33); "Team"; Tab(53); "Time"
    
    I = 0
    Do While I < CTR
        I = I + 1
        If SearchTeam = team(I) Then
            picResults.Print place(I); Tab(10); runner(I); Tab(33); team(I); Tab(53); minutes(I); ":"; seconds(I)
            found = True
        End If
    Loop
    
    If (Not found) Then
        output = MsgBox("That team did not run this race")
    End If
    
End Sub
