VERSION 5.00
Begin VB.Form frmTeamStats 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   11715
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11715
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn2 
      Caption         =   "Return to the Start Form"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16560
      TabIndex        =   8
      Top             =   9720
      Width           =   1815
   End
   Begin VB.PictureBox picDivision 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   14040
      ScaleHeight     =   2835
      ScaleWidth      =   1755
      TabIndex        =   7
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdDivisionTitles 
      BackColor       =   &H000000C0&
      Caption         =   "Click to Display the Years the Twins Won Division Titles"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      MaskColor       =   &H000000C0&
      TabIndex        =   6
      Top             =   5280
      Width           =   6495
   End
   Begin VB.PictureBox picAttendance 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   12120
      ScaleHeight     =   1035
      ScaleWidth      =   3675
      TabIndex        =   4
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton cmdDisplayYear 
      Caption         =   "Display Data"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdTeamData 
      BackColor       =   &H00800000&
      Caption         =   "Display the Team Record for Every Year in Chronological Order"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      MaskColor       =   &H00800000&
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.PictureBox picRecord 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   10695
      Left            =   2760
      ScaleHeight     =   10635
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Image imaDivisionChamps 
      Height          =   2895
      Left            =   9960
      Picture         =   "frmTeamStats.frx":0000
      Top             =   6120
      Width           =   3915
   End
   Begin VB.Label lblRecord 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Enter a year to display the fan attendance for that year."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11040
      TabIndex        =   5
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "frmTeamStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form dealing with team statistical information

Dim Pos As Integer, Pass As Integer
Dim TempSeason As Long, TempWins As Long, TempLosses As Long, TempAttendance As Single, TempChampions As String

Private Sub cmdDisplayYear_Click() 'input year from textbox and display attendance for that year
Dim InputYear As Single, Found As Boolean, C As Single

picAttendance.Cls   'clear attendance picture box

'input year from use and initialize variables
InputYear = txtYear.Text
C = 0
Found = False

Do While ((Not Found) And C < CtrTeam) 'search list for year until found or the end of the list is reached
    C = C + 1
    If Season(C) = InputYear Then   'once a match is found, drop out of the loop
        Found = True
    End If
Loop

'Print results and display error if input is not valid
If Found = True Then
    picAttendance.Print "The attendance for"; InputYear
    picAttendance.Print "was "; FormatNumber(Attendance(C), 0)
    Else: MsgBox "Error. The year you entered is not valid."
End If
   

txtYear.Text = ""   'clear textbox


End Sub

Private Sub cmdDivisionTitles_Click() 'display years the Twins won division titles
Dim L As Integer

picDivision.Cls 'clear picture box

'print heading for list
picDivision.Print "The Twins won"
picDivision.Print "their division in:"
picDivision.Print "**********************"

'display only the year in which champions = "yes"
For L = 1 To CtrTeam
    If Champions(L) = "YES" Then
        picDivision.Print Season(L)
    End If
Next L


End Sub

Private Sub cmdReturn2_Click() 'Return to the start form
    frmStart.Show
    frmTeamStats.Hide
End Sub

Private Sub cmdTeamData_Click() 'Displays team records from 1961-2010 in chronological order

picRecord.Cls 'clear picture box

'Use bubble sort to display team records in chronological order
For Pass = 1 To CtrTeam - 1         'keep track of how many passes
    For Pos = 1 To CtrTeam - Pass 'keep track of how many comparisons
        If Season(Pos) > Season(Pos + 1) Then
            TempSeason = Season(Pos) 'exchange years if out of order
            Season(Pos) = Season(Pos + 1)
            Season(Pos + 1) = TempSeason
            
            TempWins = Wins(Pos) 'exchange number of wins
            Wins(Pos) = Wins(Pos + 1)
            Wins(Pos + 1) = TempWins
            
            TempLosses = Losses(Pos) 'exchange number of losses
            Losses(Pos) = Losses(Pos + 1)
            Losses(Pos + 1) = TempLosses
            
            TempAttendance = Attendance(Pos) 'exchange attendance
            Attendance(Pos) = Attendance(Pos + 1)
            Attendance(Pos + 1) = TempAttendance
            
            TempChampions = Champions(Pos) 'exchange attendance
            Champions(Pos) = Champions(Pos + 1)
            Champions(Pos + 1) = TempChampions
            
            
        End If
    Next Pos
Next Pass

'print heading for table
picRecord.Print "Minnesota Twins: Wins and Losses Each Season"
picRecord.Print
picRecord.Print "Year"; Tab(11); "Number of Wins"; Tab(30); "Number of Losses"
picRecord.Print "**************************************************************"

'print the sorted list
For C = 1 To CtrTeam
    picRecord.Print Season(C); Tab(15); Wins(C); Tab(35); Losses(C)
Next C


End Sub
