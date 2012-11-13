VERSION 5.00
Begin VB.Form TournamentForm 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   1545
   ClientTop       =   1245
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MousePointer    =   8  'Size NW SE
   ScaleHeight     =   9615
   ScaleWidth      =   12180
   Begin VB.CommandButton cmdChron 
      BackColor       =   &H008080FF&
      Caption         =   "Arrange Chronologically"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Our Conference"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   3495
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H0080FF80&
      Caption         =   "Read Tournaments"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2160
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   6120
      ScaleHeight     =   4755
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "TournamentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WaterPoloProject
'TournamentForm
'Bobby Chapman
'Written 3/16/2009
'Objective- to read the tournmament array and sort into
'chronological order

Option Explicit
'declare global variables
Dim Place(1 To 5) As String, Dates(1 To 5) As String, Ctr As Integer

Private Sub cmdRead_Click()

'sets ctr to 0
Ctr = 0

'opens the file to be read
Open App.Path & "\Tournaments.txt" For Input As #1

'prints the header
picResults.Print "Location"; Tab(25); "Date"
picResults.Print "****************************************************"

'reads the file into 2 arrays
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Place(Ctr), Dates(Ctr)
    'prints the arrays
    picResults.Print Place(Ctr); Tab(25); Dates(Ctr)
Loop

'makes the Chronological button visible
cmdRead.Visible = False
cmdChron.Visible = True

'closes the file used for input
Close #1
End Sub

Private Sub cmdChron_Click()
'declare local variables
Dim J As Integer, Pass As Integer, Pos As Integer
Dim TempDates As String, TempPlace As String

'clears the picResults
picResults.Cls

'prints the header
picResults.Print "Location", Tab(25); "Date"
picResults.Print "***************************"

'arranges the arrays in chronological order
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Dates(Pos) < Dates(Pos + 1) Then
            TempDates = Dates(Pos)
            Dates(Pos) = Dates(Pos + 1)
            Dates(Pos + 1) = TempDates
            
            TempPlace = Place(Pos)
            Place(Pos) = Place(Pos + 1)
            Place(Pos + 1) = TempPlace
        End If
    Next Pos
Next Pass

'prints the arrays
For J = 1 To Ctr
    picResults.Print Place(J); Tab(25); Dates(J)
Next J
End Sub

Private Sub cmdBack_Click()
'goes back to Our Conference page
TournamentForm.Hide
OurConferenceForm.Show
End Sub
