VERSION 5.00
Begin VB.Form frmschedule 
   BackColor       =   &H00400040&
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   735
      Left            =   2400
      TabIndex        =   13
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdSchedule 
      Caption         =   "Load Away Schedule"
      Height          =   735
      Left            =   9480
      TabIndex        =   12
      Top             =   600
      Width           =   4695
   End
   Begin VB.PictureBox picResults 
      Height          =   6615
      Left            =   9480
      ScaleHeight     =   6555
      ScaleWidth      =   4635
      TabIndex        =   11
      Top             =   1560
      Width           =   4695
   End
   Begin VB.PictureBox picPittsburg 
      Height          =   1335
      Left            =   2520
      Picture         =   "frmschedule.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   6840
      Width           =   1935
   End
   Begin VB.PictureBox picGreenBay 
      Height          =   1335
      Left            =   4320
      Picture         =   "frmschedule.frx":0E64
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   4920
      Width           =   1935
   End
   Begin VB.PictureBox picChicago 
      Height          =   1575
      Left            =   720
      Picture         =   "frmschedule.frx":1C35
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.PictureBox picCharlotte 
      Height          =   1575
      Left            =   720
      Picture         =   "frmschedule.frx":2D16
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   4920
      Width           =   1815
   End
   Begin VB.PictureBox picArizona 
      Height          =   1455
      Left            =   4320
      Picture         =   "frmschedule.frx":3FE3
      ScaleHeight     =   1395
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblPittsburg 
      BackColor       =   &H00400040&
      Caption         =   "5: Pittsburg"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label lblGreenBay 
      BackColor       =   &H00400040&
      Caption         =   "4: Green Bay"
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
      Left            =   6600
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblChicago 
      BackColor       =   &H00400040&
      Caption         =   "1: Chicago"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblcharlotte 
      BackColor       =   &H00400040&
      Caption         =   "2: Charlotte"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label lblArizona 
      BackColor       =   &H00400040&
      Caption         =   "3: Arizona"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblschedule 
      BackColor       =   &H00400040&
      Caption         =   $"frmschedule.frx":53AE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmschedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Brett Favre Fan Club
'Form Name: frmschedule
'Author: Kory LaCroix
'Date Written: 10/19/08
'Objective: To select a city and away game to attend to watch Brett Favre and the Vikings
Option Explicit

Private Sub cmdChoose_Click()
'The following ask the user to enter a number which is connected to a city
'the program then gives the user the cost of traveling to the city and the cost of a ticket
'it also then moves on to the next form
choice = InputBox("Please enter the number of the city that you would like to travel to.")
    If choice = 1 Then
        MsgBox ("A game ticket and a flight to Chicago will cost you $180.00")
        runningtotal = runningtotal + 180
        frmhotels.Show
        frmschedule.Hide
    ElseIf choice = 2 Then
        MsgBox ("A game ticket and a flight to Charlotte will cost you $375.00")
        runningtotal = runningtotal + 375
        frmhotels.Show
        frmschedule.Hide
    ElseIf choice = 3 Then
        MsgBox ("A game ticket and a flight to Arizona will cost you $500.00")
        runningtotal = runningtotal + 500
        frmhotels.Show
        frmschedule.Hide
    ElseIf choice = 5 Then
        MsgBox ("A game ticket and a flight to Pittsburg will cost you $525.00")
        runningtotal = runningtotal + 525
        frmhotels.Show
        frmschedule.Hide
    ElseIf choice = 4 Then
        MsgBox ("A game ticket and a flight to Packerland will cost you $350.00")
        runningtotal = runningtotal + 350
        frmhotels.Show
        frmschedule.Hide
    Else
        MsgBox ("You did not enter a number 1 through 5, please try again.")
    End If
    
End Sub

Private Sub cmdSchedule_Click()
Dim dates(1 To 50) As String
Dim location(1 To 50) As String
Dim time(1 To 50) As String

'this opens up a file containg the reaming vikings schedule
Open App.Path & "\schedule.txt" For Input As #2

Do While Not EOF(2)
    CTR = CTR + 1
    Input #2, dates(CTR), location(CTR), time(CTR)
Loop

picResults.Print "Date"; Tab(25); "Location"; "   Time"
picResults.Print "*******************************************"

For j = 1 To CTR
    picResults.Print dates(j); Tab(25); location(j); "   "; time(j)
Next j

cmdSchedule.Visible = False
End Sub

