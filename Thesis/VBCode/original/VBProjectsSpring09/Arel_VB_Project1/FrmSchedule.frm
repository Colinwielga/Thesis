VERSION 5.00
Begin VB.Form FrmSchedule 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Click Here To Reserve Your Spot For A Home Game This Season!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7200
      Width           =   3615
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   9
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      Picture         =   "FrmSchedule.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton CmdMain 
      BackColor       =   &H8000000E&
      Caption         =   "Go Back to Main Menu"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   9
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Picture         =   "FrmSchedule.frx":16CE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000D&
      Caption         =   "Go!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H8000000D&
      Caption         =   "Go!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000E&
      Caption         =   "Click to Begin!"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "Go!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox txtday 
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   6360
      Width           =   3015
   End
   Begin VB.TextBox txtmonth 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox txtTeam 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      Height          =   5895
      Left            =   4440
      ScaleHeight     =   5835
      ScaleWidth      =   5355
      TabIndex        =   2
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By Weekday - Monday = Sunday"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   5880
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By Month"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Welcome to the Home Game Viewer!"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   36
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   1215
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   9375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By Opponent (Mascot - Twins, white sox, Angels... etc)"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   3615
   End
End
Attribute VB_Name = "FrmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Title: Minnesota Twins Fan
'Form Name: FrmSchedule
'Project By: Stephanie Arel
'Date Written: 3/16/2009
'The purpose of this form is to allow the user to see the twins schedule by opponent, month and date.

Dim hmteam(1 To 100) As String
Dim hmday(1 To 100) As String
Dim hmmonth(1 To 100) As String
Dim hmdate(1 To 100) As Single

Dim I As Integer
Dim hmCtr As Integer


Dim J As Integer

Option Explicit

Private Sub CmdMain_Click()
'Takes the user back to the main menu
FrmSchedule.Hide
FrmMain.Show

End Sub

Private Sub CmdQuit_Click()
'Ends the program
End

End Sub

Private Sub Command1_Click()
'Takes the user to purchase tickets.
FrmSchedule.Hide
FrmTickets.Show

End Sub

Private Sub Command3_Click()
'Searches through the list by day.
Dim day As String
Dim Found As Boolean
Found = False
picResults.Cls

day = txtday.Text

picResults.Print Tab(5); "Games on "; day; "s"
picResults.ForeColor = vbRed
picResults.Print

picResults.Print Tab(10); "Home Games"
picResults.ForeColor = vbBlack
picResults.Print "*******************************************"
picResults.Print "Opponent"; Tab(15); "Month"; Tab(30); "Date"
picResults.Print "*******************************************"

For J = 1 To hmCtr
    If day = hmday(J) Then
    Found = True
    picResults.Print hmteam(J); Tab(15); hmmonth(J); Tab(30); hmdate(J)
    End If
Next J

If Not Found Then
    MsgBox "Sorry! Please enter a valid weekday!"
End If
End Sub

Private Sub Command4_Click()

'Loads the home schedule.
Open App.Path & "\Home.txt" For Input As #1

hmCtr = 0
Do While Not EOF(1)
    hmCtr = hmCtr + 1
    Input #1, hmteam(hmCtr), hmday(hmCtr), hmmonth(hmCtr), hmdate(hmCtr)
Loop


End Sub


Private Sub Command5_Click()
'Searches through the list by month.
Dim Month As String
Dim Found As Boolean
Found = False
picResults.Cls

Month = txtmonth.Text

picResults.Print Tab(5); "Games in "; Month
picResults.ForeColor = vbRed
picResults.Print

picResults.Print Tab(10); "Home Games"
picResults.ForeColor = vbBlack
picResults.Print "*******************************************"

For J = 1 To hmCtr
    If Month = hmmonth(J) Then
    Found = True
    picResults.Print hmteam(J); Tab(15); hmday(J); Tab(40); hmmonth(J); Tab(45); hmdate(J)
    End If
Next J

If Not Found Then
    MsgBox "Sorry! The Twin's Don't Play this month! Please enter a valid month!"
End If

End Sub

Private Sub Command8_Click()
'Searches through the list by opponent.

Dim team As String
Dim Found As Boolean
Found = False
picResults.Cls

team = txtTeam.Text

picResults.Print Tab(5); "Twins vs. The "; team
picResults.ForeColor = vbRed
picResults.Print

picResults.Print Tab(10); "Home Games"
picResults.ForeColor = vbBlack
picResults.Print "********************************************"

For J = 1 To hmCtr
    If team = hmteam(J) Then
        Found = True
        picResults.Print hmteam(J); Tab(15); hmday(J); Tab(26); hmmonth(J); Tab(40); hmdate(J)
    End If
Next J

If Not Found Then
    picResults.ForeColor = vbRed
    picResults.Print "Sorry! Either you have typed an invalid team, "
    picResults.Print "or they do not play any home games against the Twins."
End If


End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
