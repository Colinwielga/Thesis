VERSION 5.00
Begin VB.Form DifferentTeamsForm 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   9630
   ClientLeft      =   1650
   ClientTop       =   1125
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   9630
   ScaleWidth      =   12045
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
      Height          =   1335
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Search team location"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H0080FF80&
      Caption         =   "Read Teams"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2715
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblMonmonth 
      Caption         =   "Monmonth Crab People"
      BeginProperty Font 
         Name            =   "DotumChe"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "DifferentTeamsForm.frx":0000
      Top             =   3000
      Width           =   9060
   End
End
Attribute VB_Name = "DifferentTeamsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WaterPoloProject
'DifferentTeamsForm
'Bobby Chapman
'Written 3/16/2009
'Objective- To list all the teams in our Conference and search for a team location

Option Explicit
'declare global variables
Dim TeamName(1 To 10) As String, Location(1 To 10) As String
Dim Ctr As Integer

Private Sub cmdRead_Click()
'declare local variables
Dim I As Integer

'sets ctr to 0
Ctr = 0

'opens the file to be read
Open App.Path & "\ConferenceTeams.txt" For Input As #1

'prints the header
picResults.Print "Team Name"
picResults.Print "*******************"

'reads the file into 2 arrays
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, TeamName(Ctr), Location(Ctr)
Loop

'prints the 2 arrays
For I = 1 To Ctr
    picResults.Print TeamName(I)
Next I

'makes the Search button visible
cmdSearch.Visible = True
cmdRead.Visible = False

'closes the input file
Close #1
End Sub

Private Sub cmdSearch_Click()
'declare local variables
Dim Search As String, Found As Boolean, J As Single

'sets J to 0
J = 0
'sets Found to false to be used in the search
Found = False

'uses an input box to search for a team location
Search = InputBox("Enter the team whose location you wish to know", "Search")

'searches using match and stop to find the searched team
Do While ((Not Found) And (J < Ctr))
    J = J + 1
    If Search = TeamName(J) Then
        Found = True
        'if the team is found, it displays a message box
        MsgBox "The team " & Search & " is located in " & Location(J), , "Team Found"
    End If
Loop

'if not found, it displays a message box
If (Not Found) Then
    MsgBox "Team not found.", , "Alert"
End If

End Sub

Private Sub cmdBack_Click()
'goes back to the OurConferenceForm
DifferentTeamsForm.Hide
OurConferenceForm.Show
End Sub
