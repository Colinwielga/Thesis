VERSION 5.00
Begin VB.Form ResultsForm 
   BackColor       =   &H00800000&
   Caption         =   "Race Results for ""Viessmann"" FIS World Cup Cross-Country"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2715
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   1215
      Left            =   600
      TabIndex        =   3
      Top             =   5760
      Width           =   2895
   End
   Begin VB.CommandButton SelCountry 
      Caption         =   "Select Country"
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton DisResults 
      Caption         =   "Display Race Results"
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.PictureBox DisplayResults 
      Height          =   10815
      Left            =   5400
      ScaleHeight     =   10755
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "ResultsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Skier(1 To 72) As String, Country(1 To 72) As String, Time(1 To 72) As String
Dim Place(1 To 72) As Integer, Minutes(1 To 72) As Integer, Seconds(1 To 72) As Single
Public PATH As String
Private Sub DisResults_Click()
Dim J As Integer
DisplayResults.Cls
Open PATH & "RaceResults.txt" For Input As #1
For J = 1 To 72
    Input #1, Place(J), Skier(J), Country(J), Time(J), Minutes(J)
Next J
DisplayResults.Print "Skier"; Tab(30); "Country"; Tab(40); "Time"
For J = 1 To 72
    DisplayResults.Print Place(J); Skier(J); Tab(30); Country(J); Tab(40); Time(J)
Next J
Close #1
SelCountry.Enabled = True
End Sub

Private Sub Form_Load()
PATH = "N:\CS130\handin\ChrisHenkel\"
SelCountry.Enabled = False
Picture1.Picture = LoadPicture(PATH & "FREEMAN_JUSTIN.jpg")
End Sub

Private Sub Quit_Click()
End
End Sub

Private Sub SelCountry_Click()
ResultsForm.Hide
TeamForm.Show
End Sub
