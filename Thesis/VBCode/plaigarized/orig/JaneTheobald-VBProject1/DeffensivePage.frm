VERSION 5.00
Begin VB.Form frmDeffensivePage 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   11760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17145
   LinkTopic       =   "Form1"
   Picture         =   "DeffensivePage.frx":0000
   ScaleHeight     =   11760
   ScaleWidth      =   17145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10200
      Width           =   1815
   End
   Begin VB.CommandButton cmdReadStats 
      BackColor       =   &H0000FFFF&
      Caption         =   "Read Stats"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtDescription 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "DeffensivePage.frx":2A428A
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   14280
      Picture         =   "DeffensivePage.frx":2A4305
      Top             =   10080
      Width           =   1680
   End
End
Attribute VB_Name = "frmDeffensivePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Page for defensive stats - search for player by name and then calculate
'Fielding percentage to a msgbox

Private Sub cmdReadStats_Click()
'declare varilables
Dim Found As Boolean
Dim Name As String
Dim N As Integer
Dim FPA As Integer, FPB As Integer
Dim temporary As Integer
'intialize counter
Counter = 0
Counter = Counter + 1

'Read in file
Open App.Path & "\DefensiveStats.txt" For Input As #1
Do While Not EOF(1)
    Counter = Counter + 1
    Input #1, PlayerNumber(Counter), PlayerName(Counter), PutOuts(Counter), Assits(Counter), Errors(Counter)
Loop
'initilizale counter/define variables
Name = InputBox("Data was read successfully, enter player name you want to find:")
Found = False
    FPA = PutOuts(Counter) + Assits(Counter)
    FPB = (PutOuts(Counter) + Assits(Counter) + Errors(Counter))
    FieldingPercentage(Counter) = FPA / FPB
'data was successfully read - use input box to search for players
'Loop Through Names and print results
    Do While ((Not Found) And (N < Counter))
    N = N + 1

        If Name = PlayerName(N) Then
            Found = True
            temporary = N
        End If
    Loop
'if player is found print
If Found = True Then
        MsgBox PlayerName(temporary) & " has " & PutOuts(temporary) & " putouts, " & Assits(temporary) & " assits, and " & Errors(temporary) & " errors on the season. " & PlayerName(temporary) & "'s fielding percentage on the year is " & FormatPercent(FieldingPercentage(temporary)) & "."
'if error print this:
    Else: Found = False
        'name not found results
        MsgBox "Please reenter name, player not found."
    End If
Close #1
End Sub
'return to main page
Private Sub cmdReturn_Click()
frmAHomePage.Show
frmDeffensivePage.Hide
frmPitchingPage.Hide
frmOffensivePage.Hide
End Sub

Private Sub Image1_Click()
MsgBox "Big De - No E!"
End Sub
