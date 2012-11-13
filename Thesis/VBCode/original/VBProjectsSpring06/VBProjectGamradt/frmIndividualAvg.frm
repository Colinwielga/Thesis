VERSION 5.00
Begin VB.Form frmIndividualAvg 
   BackColor       =   &H00808000&
   Caption         =   "Calculate Individual Average"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Height          =   4095
      Left            =   6600
      ScaleHeight     =   4035
      ScaleWidth      =   4275
      TabIndex        =   19
      Top             =   2400
      Width           =   4335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Averages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   14
      Top             =   6000
      Width           =   1695
   End
   Begin VB.PictureBox picAvgTurnovers 
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   13
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cmdInputTurnovers 
      Caption         =   "Turnovers"
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   5160
      Width           =   1695
   End
   Begin VB.PictureBox picAvgRebounds 
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   11
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdInputRebounds 
      Caption         =   "Rebounds"
      Height          =   615
      Left            =   3960
      TabIndex        =   10
      Top             =   4320
      Width           =   1695
   End
   Begin VB.PictureBox picAvgMinutes 
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   9
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdInputMinutes 
      Caption         =   "Minutes"
      Height          =   615
      Left            =   3960
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.PictureBox picAvgPoints 
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.FileListBox fileTeam 
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   540
      Left            =   120
      Pattern         =   "*Names.txt*"
      TabIndex        =   6
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Display Team Roster"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   720
      ScaleHeight     =   2475
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdInputPoints 
      Caption         =   "Points"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdNavigatePlayerMatch 
      Caption         =   "Player Matchup"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavigateMainMenu 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   9600
      TabIndex        =   0
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label lblSavedAvg 
      BackColor       =   &H00808000&
      Caption         =   "Saved Averages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7560
      TabIndex        =   20
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblPlayerMatchup 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Individual Averages"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   555
      Left            =   2760
      TabIndex        =   18
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label lblTeamFile 
      BackColor       =   &H00808000&
      Caption         =   "*Must Highlight a Team File Before Opening*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label lblDesigner 
      BackColor       =   &H00808000&
      Caption         =   "By: Erik Gamradt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Team Manager Pro"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmIndividualAvg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Team Manager Pro (ErikGamradtVBProject.vbp)
'frmIndividualAvg (frmIndividualAvg.frm)
'Designed By: Erik Gamradt
'18 March 2006
'This form allows users to select a team from the file list, open the roster and display it, enter in game stats and find averages for points, minutes, rebounds, and turnovers, and save them to the appropriate team file.
Option Explicit
    Dim AvgPoints, AvgMinutes, AvgRebounds, AvgTurnovers As Single
    Dim I As Integer
    Dim Names As String
    Dim HomeNames(1 To 100), OutFile As String
    Dim Size As Integer
    
Private Sub cmdInputMinutes_Click()
    Dim Sum, counter As Integer
    Sum = 0
    counter = 0
    AvgMinutes = 0
    AvgMinutes = InputBox("Enter Individual Game Minute Totals (-1 to exit): ", "Minutes")
    Do While AvgMinutes <> -1  'allows for an exit when all the statistics have been entered
        Sum = Sum + AvgMinutes
        counter = counter + 1
        AvgMinutes = InputBox("Enter Individual Game Minute Totals (-1 to exit): ", "Minutes")
    Loop
    AvgMinutes = Sum / counter 'total divided by how many times you entered in data to find average
    picAvgMinutes.Print FormatNumber(AvgMinutes); "MPG"
End Sub

Private Sub cmdInputPoints_Click()
    Dim Pos As Integer
    Dim Found As Boolean
    Dim SHomeNames(1 To 100) As String
    Dim Sum, counter As Integer
    Found = False
    Pos = 0
    Sum = 0
    counter = 0
    AvgPoints = 0
    Names = InputBox("Enter Player Name from List: ", "Player Name")
    Do While Found = False And Pos < HomeNamesSize   'Test to see if name matches a player on team roster
        Pos = Pos + 1
        If Names = HomeNames(Pos) Then
            Found = True
        End If
    Loop
    If Found = False Then
        MsgBox "Invalid Player Entry for this Team.  Please try again.", vbCritical, "Invalid Player" 'gives appropiate message to user
    Else
        Pos = 0
        Found = False
        OutFile = Replace(fileTeam.FileName, "Names", "")
        Open fileTeam.Path & "\" & OutFile For Input As #2
        Do Until EOF(2)
            Pos = Pos + 1
            Input #2, SHomeNames(Pos)
        Loop
        Close #2
        Pos = 0
        Do While Found = False And Pos < HomeNamesSize  'Test to see if player already has statistical data
            Pos = Pos + 1
            If Names = SHomeNames(Pos) Then
                Found = True
            End If
        Loop
        If Found = True Then
            MsgBox "This player already has statistics recorded", vbCritical, "Error"
        Else
            AvgPoints = InputBox("Enter Individual Game Point Totals (-1 to exit): ", "Points")
            Do While AvgPoints <> -1
                Sum = Sum + AvgPoints
                counter = counter + 1
                AvgPoints = InputBox("Enter Individual Game Point Totals (-1 to exit): ", "Points")
            Loop
            AvgPoints = Sum / counter
            picAvgPoints.Print FormatNumber(AvgPoints); "PPG"
        End If
    End If
End Sub

Private Sub cmdInputRebounds_Click()
    Dim Sum, counter As Integer
    Sum = 0
    counter = 0
    AvgRebounds = 0
    AvgRebounds = InputBox("Enter Individual Game Rebound Totals (-1 to exit): ", "Rebounds")
    Do Until AvgRebounds = -1
        Sum = Sum + AvgRebounds
        counter = counter + 1
        AvgRebounds = InputBox("Enter Individual Game Rebound Totals (-1 to exit): ", "Rebounds")
    Loop
    AvgRebounds = Sum / counter
    picAvgRebounds.Print FormatNumber(AvgRebounds); "RPG"
End Sub

Private Sub cmdInputTurnovers_Click()
    Dim Sum, counter As Integer
    Sum = 0
    counter = 0
    AvgTurnovers = 0
    AvgTurnovers = InputBox("Enter Individual Game Turnovers Totals (-1 to exit): ", "Turnovers")
    Do While AvgTurnovers <> -1
        Sum = Sum + AvgTurnovers
        counter = counter + 1
        AvgTurnovers = InputBox("Enter Individual Game Turnovers Totals (-1 to exit): ", "Turnovers")
    Loop
    AvgTurnovers = Sum / counter
    picAvgTurnovers.Print FormatNumber(AvgTurnovers); "TPG"
End Sub

Private Sub cmdNavigateMainMenu_Click()
    frmMainMenu.Show
    frmIndividualAvg.Hide
End Sub

Private Sub cmdNavigatePlayerMatch_Click()
    frmPlayerMatch.Show
    frmIndividualAvg.Hide
End Sub

Private Sub cmdOpen_Click()
    Dim Pos As Integer
    picResults.Cls
    Pos = 0
    Open fileTeam.Path & "\" & fileTeam.FileName For Input As #1 'opens file that is selected from the file list
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, HomeNames(Pos)
        picResults.Print HomeNames(Pos)
    Loop
    HomeNamesSize = Pos
    Close #1
End Sub
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSave_Click()
    Dim Pos As Integer
    picAvgTurnovers.Cls  'when saved, the averages are erased for next player entry
    picAvgRebounds.Cls
    picAvgMinutes.Cls
    picAvgPoints.Cls
    Pos = 0
    OutFile = Replace(fileTeam.FileName, "Names", "")  'changes the selected file name to another file so new data can be saved to it
    Open fileTeam.Path & "\" & OutFile For Append As #1  'opens file for writing
    Write #1, Names, FormatNumber(AvgPoints), FormatNumber(AvgMinutes), FormatNumber(AvgRebounds), FormatNumber(AvgTurnovers)
    Close #1
    picResults2.Print Names; Tab(15); FormatNumber(AvgPoints); " PPG"; Tab(26); FormatNumber(AvgMinutes); " MPG"; Tab(37); FormatNumber(AvgRebounds); " RPG"; Tab(49); FormatNumber(AvgTurnovers); " TPG"
End Sub

