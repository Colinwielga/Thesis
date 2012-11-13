VERSION 5.00
Begin VB.Form Basketball_Jobs 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   FillColor       =   &H80000005&
   ForeColor       =   &H80000003&
   LinkTopic       =   "Form1"
   Picture         =   "bball_jobs.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Message 
      Caption         =   "Click here last"
      Height          =   975
      Left            =   6360
      TabIndex        =   26
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton years_make_equal 
      Caption         =   "Whole year comparison"
      Height          =   1095
      Left            =   6360
      TabIndex        =   24
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton load_file 
      Caption         =   "Click to load file"
      Height          =   975
      Left            =   6360
      TabIndex        =   23
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox results 
      Height          =   1695
      Left            =   5880
      ScaleHeight     =   1635
      ScaleWidth      =   3435
      TabIndex        =   22
      Top             =   5400
      Width           =   3495
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   3360
      TabIndex        =   21
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox job 
      Height          =   2415
      Left            =   3360
      TabIndex        =   19
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox player 
      Height          =   1935
      Left            =   3480
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton week_to_year 
      Caption         =   "Week to year comparison"
      Height          =   1095
      Left            =   6360
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label18 
      Caption         =   "Select the number of the job you would most like to have"
      Height          =   495
      Left            =   3360
      TabIndex        =   20
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label17 
      Caption         =   "12. Veterinarian"
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label16 
      Caption         =   "11. Teacher"
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "10. Physician"
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "9. Optometrist"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "8. Mathematician"
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "7. Lawyer"
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "6. Insurance Agent"
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "5. Dentist"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "4. Computer Programmer"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "3. Chiropractor"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "2. Architect"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "1. Accountant  "
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Select the number of your favorite player "
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "4. Allen Iverson"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "3. Kobe Bryant"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "2. Shaquille O'Neal"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "1. Kevin Garnett"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Basketball_Jobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Path As String

Dim nba_player(1 To 4) As String
Dim nba_salary(1 To 4) As Double
Dim jobs(1 To 12) As String
Dim jobs_salary(1 To 12) As Double
Dim years As Integer
Dim salary_in_week As Double
Dim years_equal As Integer
Dim I As Double
Dim J As Double
Dim NBA As String
Dim professions As String



Private Sub Form_Load()
Path = "M:\CS130\Projects\"

End Sub


Private Sub load_file_Click()
Open Path & "project_basketball.txt" For Input As #1

For I = 1 To 4
    Input #1, nba_player(I), nba_salary(I)
Next I


Open Path & "project_jobs.txt" For Input As #2

For J = 1 To 12
    Input #2, jobs(J), jobs_salary(J)
Next J

End Sub

Private Sub Message_Click()
    MsgBox "Crazy isn't it? :)", , "Message"

End Sub

Private Sub Quit_Click()
End
End Sub


Private Sub week_to_year_Click()


NBA = player.Text
professions = job.Text


salary_in_week = nba_salary(NBA) / 52
years_equal = salary_in_week / jobs_salary(professions)

results.Print "It would take you"; years_equal; "years to make as much as"; Tab(1); nba_player(NBA); " makes in one week"
End Sub

Private Sub years_make_equal_Click()
NBA = player.Text
professions = job.Text


years = nba_salary(NBA) / jobs_salary(professions)

results.Print "It would take you"; years; "years to make as much as"; Tab(1); nba_player(NBA); " makes in one year"

End Sub
