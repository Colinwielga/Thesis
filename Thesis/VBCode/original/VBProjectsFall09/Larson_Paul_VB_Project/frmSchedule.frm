VERSION 5.00
Begin VB.Form frmSchedule 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   13845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16200
   LinkTopic       =   "Form1"
   ScaleHeight     =   13845
   ScaleWidth      =   16200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdView2 
      Caption         =   "Show Schedule"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   600
      TabIndex        =   4
      Top             =   12000
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11520
      TabIndex        =   3
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Go Back to Home Page"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10320
      TabIndex        =   2
      Top             =   12360
      Width           =   4695
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Click to See Schedule"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   1
      Top             =   12240
      Width           =   8655
   End
   Begin VB.PictureBox picResults 
      Height          =   11175
      Left            =   600
      ScaleHeight     =   11115
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   360
      Width           =   9975
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dates(1 To 20) As String, Opponent(1 To 20) As String, City(1 To 20) As String, State(1 To 20) As String, Time(1 To 20) As String, CTR As Integer 'declaring variables
Private Sub cmdQuit_Click()
    End 'ending the program
End Sub

Private Sub cmdReturn_Click()
    frmHome.Show 'show the home page
    frmSchedule.Hide 'hiding the schedule
    
End Sub

Private Sub cmdView_Click()
    CTR = 0 'setting the counter to zero
    picResults.Print "Date", Tab(15); "Opponent/Meet", Tab(75); "City", Tab(95); "State", Tab(115); "Time" 'picresults
    picResults.Print "********************************************************************************************************************************************************************"
    
    Open App.Path & "\Schedule.txt" For Input As #1 'opening a file for schedule
    Do While Not EOF(1)
        CTR = CTR + 1 'going through the file with the counter
        Input #1, Dates(CTR), Opponent(CTR), City(CTR), State(CTR), Time(CTR) 'setting rows for all parts of the schedule
        picResults.Print Dates(CTR), Tab(15); Opponent(CTR), Tab(75); City(CTR), Tab(95); State(CTR), Tab(115); Time(CTR); picResults
        picResults.Print
    Loop
    cmdView.Visible = False
    cmdView2.Visible = True
    
End Sub

Private Sub cmdView2_Click()
    picResults.Cls 'clearing the list when clicked for a second time
    picResults.Print "Date", Tab(15); "Opponent/Meet", Tab(75); "City", Tab(95); "State", Tab(115); "Time"; picResults
    picResults.Print "***********************************************************************************************************************************************************************"
    Dim J As Integer 'declaring variables
    J = 0
    Do While J < CTR 'array for the schedule
        J = J + 1
        picResults.Print Dates(J), Tab(15); Opponent(J), Tab(75); City(J), Tab(95); State(J), Tab(115); Time(J)
        picResults.Print
    Loop 'ending the array
End Sub
