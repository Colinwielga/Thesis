VERSION 5.00
Begin VB.Form TeamForm 
   BackColor       =   &H000000C0&
   Caption         =   "Select Team"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Return 
      Caption         =   "Return to First Page"
      Height          =   1215
      Left            =   5400
      TabIndex        =   8
      Top             =   8760
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000C0&
      Height          =   2655
      Left            =   480
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   480
      TabIndex        =   6
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton TimeAve 
      Caption         =   "Find Average Time for the Selected Team."
      Height          =   975
      Left            =   480
      TabIndex        =   5
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton TeamAve 
      Caption         =   "Find Average Position of Selected Team."
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.PictureBox Results 
      BackColor       =   &H00FFFFFF&
      Height          =   6255
      Left            =   4800
      ScaleHeight     =   6195
      ScaleWidth      =   5715
      TabIndex        =   3
      Top             =   1680
      Width           =   5775
   End
   Begin VB.CommandButton Search 
      Caption         =   "Search"
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox TeamSelect 
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "Enter team you wish to search for; use 3 letter country code. Eg. USA for the United States."
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "TeamForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim J As Integer
Dim Skier(1 To 72) As String, Country(1 To 72) As String, Time(1 To 72) As String
Dim Place(1 To 72) As Integer, Minutes(1 To 72) As Single
Private Sub Form_Load()
Open ResultsForm.PATH & "RaceResults.txt" For Input As #1
For J = 1 To 72 'load skiers
     Input #1, Place(J), Skier(J), Country(J), Time(J), Minutes(J)
Next J
Close #1
TeamAve.Enabled = False
TimeAve.Enabled = False
Picture1.Picture = LoadPicture(ResultsForm.PATH & "SWENSON_CARL.jpg")
End Sub

Private Sub Quit_Click()
End
End Sub

Private Sub Return_Click()
TeamForm.Hide
ResultsForm.Show
End Sub

Private Sub Search_Click()
 'search for team
Results.Cls
Results.Print "Skier"; Tab(25); "Country"; Tab(40); "Time"
For J = 1 To 72
    If TeamSelect = Country(J) Then 'select the country's skiers
           Results.Print Place(J); Skier(J); Tab(25); Country(J); Tab(40); Time(J)
    End If
Next J
TeamAve.Enabled = True
TimeAve.Enabled = True
End Sub

Private Sub TeamAve_Click()
Dim CTR As Integer, TempPlace As Integer, AvPlace As Single
CTR = 0
TempPlace = 0
AvPlace = 0
For J = 1 To 72 'Get the numerator and denomator for finding the average
    If TeamSelect = Country(J) Then
        TempPlace = TempPlace + Place(J)
        CTR = CTR + 1
    End If
Next J
AvPlace = TempPlace / CTR 'find the avereage
Results.Print "Team's Average position:"; Round(AvPlace); "out of 72."
End Sub

Private Sub TimeAve_Click()
Dim CTR As Integer, TeamTime As Single, AvTime As Double
CTR = 0
For J = 1 To 72
    If TeamSelect = Country(J) Then 'get all the stuff you need for finding the average time
        CTR = CTR + 1
        TeamTime = TeamTime + Minutes(J)
    End If
Next J
AvTime = TeamTime / CTR 'find average time
Results.Print "Average Team Time:"; Round(AvTime); "minutes."
End Sub
