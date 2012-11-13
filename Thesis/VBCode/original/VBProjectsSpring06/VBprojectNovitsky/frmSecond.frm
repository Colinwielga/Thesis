VERSION 5.00
Begin VB.Form frmSecond 
   BackColor       =   &H000000FF&
   Caption         =   "The calculations"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   8985
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdScheduel 
      Caption         =   "Check Your Schedule!"
      Height          =   1215
      Left            =   5280
      TabIndex        =   9
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdPromo 
      Caption         =   "WINNING WITH SPORTS CALCULATOR!"
      Height          =   1215
      Left            =   1680
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdLocal 
      Caption         =   "Local Team Records"
      Height          =   1215
      Index           =   5
      Left            =   4320
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdContacts 
      Caption         =   "Contacts"
      Height          =   1215
      Index           =   3
      Left            =   2640
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H8000000D&
      Caption         =   "Top 5 Information! (The more you know!)"
      Height          =   1215
      Index           =   2
      Left            =   5280
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSports 
      Caption         =   "Sports Options"
      Height          =   1215
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdMatches 
      Caption         =   "Matches"
      Height          =   1215
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label lblOptions 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Options are:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   1920
      TabIndex        =   8
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label lblDisplay 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   720
      TabIndex        =   7
      Top             =   0
      Width           =   8295
   End
   Begin VB.Image imgSecond 
      Height          =   7290
      Left            =   0
      Picture         =   "frmSecond.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmSecond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdContacts_Click(Index As Integer) 'opens form contacts
    frmContact.Show
    frmSecond.Hide
End Sub

Private Sub cmdInfo_Click(Index As Integer) 'opens form info
    frmClubs.Show
    frmSecond.Hide
End Sub

Private Sub cmdLocal_Click(Index As Integer) 'opens form local
    frmSecond.Hide
    frmLocal.Show
End Sub

Private Sub cmdMatches_Click(Index As Integer) 'opens form matches
    frmSecond.Hide
    frmMatches.Show
    frmMatches.Cls 'displays the matches and score in order from 1st to 14th based on score
    frmMatches.Print "Your Matches and Scores in Order From Best to Worst Are:"
    frmMatches.Print "1. " & rankname(1) & "  " & FormatNumber(total(1))
    frmMatches.Print "2. " & rankname(2) & "  " & FormatNumber(total(2))
    frmMatches.Print "3. " & rankname(3) & "  " & FormatNumber(total(3))
    frmMatches.Print "4. " & rankname(4) & "  " & FormatNumber(total(4))
    frmMatches.Print "5. " & rankname(5) & "  " & FormatNumber(total(5))
    frmMatches.Print "6. " & rankname(6) & "  " & FormatNumber(total(6))
    frmMatches.Print "7. " & rankname(7) & "  " & FormatNumber(total(7))
    frmMatches.Print "8. " & rankname(8) & "  " & FormatNumber(total(8))
    frmMatches.Print "9. " & rankname(9) & "  " & FormatNumber(total(9))
    frmMatches.Print "10. " & rankname(10) & "  " & FormatNumber(total(10))
    frmMatches.Print "11. " & rankname(11) & "  " & FormatNumber(total(11))
    frmMatches.Print "12. " & rankname(12) & "  " & FormatNumber(total(12))
    frmMatches.Print "13. " & rankname(13) & "  " & FormatNumber(total(13))
    frmMatches.Print "14. " & rankname(14) & "  " & FormatNumber(total(14))
End Sub

Private Sub cmdPromo_Click() 'opens frmwinning
    frmWinning.Show
    frmSecond.Hide
End Sub

Private Sub cmdQuit_Click() 'Ends program
    End
End Sub

Private Sub cmdScheduel_Click() ' loads frmsched and diplays a constant string
    frmSched.Show
    frmSecond.Hide
    frmSched.Print "click The times that work for you!"
End Sub

Private Sub cmdSports_Click(Index As Integer) 'opens frmoptions and displays arrays with corresponding data
    frmOptions.Show
    frmSecond.Hide
    frmOptions.Cls
    frmOptions.Print "Sport", , "forty speed", "bench weight", "shuttle time", "verticle Leap", "coordination score   "; "mile time"
    frmOptions.Print arraysport(1), , arrayforty(1), arraybench(1), arrayshuttle(1), arrayleap(1), arrayhandeye(1), "        " & arraymile(1)
    frmOptions.Print arraysport(2), , arrayforty(2), arraybench(2), arrayshuttle(2), arrayleap(2), arrayhandeye(2), "        " & arraymile(2)
    frmOptions.Print arraysport(3), arrayforty(3), arraybench(3), arrayshuttle(3), arrayleap(3), arrayhandeye(3), "        " & arraymile(3)
    frmOptions.Print arraysport(4), , arrayforty(4), arraybench(4), arrayshuttle(4), arrayleap(4), arrayhandeye(4), "        " & arraymile(4)
    frmOptions.Print arraysport(5), , arrayforty(5), arraybench(5), arrayshuttle(5), arrayleap(5), arrayhandeye(5), "        " & arraymile(5)
    frmOptions.Print arraysport(6), arrayforty(6), arraybench(6), arrayshuttle(6), arrayleap(6), arrayhandeye(6), "        " & arraymile(6)
    frmOptions.Print arraysport(7), arrayforty(7), arraybench(7), arrayshuttle(7), arrayleap(7), arrayhandeye(7), "        " & arraymile(7)
    frmOptions.Print arraysport(8), , arrayforty(8), arraybench(8), arrayshuttle(8), arrayleap(8), arrayhandeye(8), "        " & arraymile(8)
    frmOptions.Print arraysport(9), arrayforty(9), arraybench(9), arrayshuttle(9), arrayleap(9), arrayhandeye(9), "        " & arraymile(9)
    frmOptions.Print arraysport(10), , arrayforty(10), arraybench(10), arrayshuttle(10), arrayleap(10), arrayhandeye(10), "        " & arraymile(10)
    frmOptions.Print arraysport(11), , arrayforty(11), arraybench(11), arrayshuttle(11), arrayleap(11), arrayhandeye(11), "        " & arraymile(11)
    frmOptions.Print arraysport(12), , arrayforty(12), arraybench(12), arrayshuttle(12), arrayleap(12), arrayhandeye(12), "        " & arraymile(12)
    frmOptions.Print arraysport(13), , arrayforty(13), arraybench(13), arrayshuttle(13), arrayleap(13), arrayhandeye(13), "        " & arraymile(13)
    frmOptions.Print arraysport(14), , arrayforty(14), arraybench(14), arrayshuttle(14), arrayleap(14), arrayhandeye(14), "        " & arraymile(14)
    
End Sub

Private Sub Form_Load() 'puts persons inputed name as a lable, to personalize the fun!
    lblDisplay.Caption = person
End Sub

