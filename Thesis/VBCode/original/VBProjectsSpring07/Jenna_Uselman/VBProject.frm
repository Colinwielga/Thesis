VERSION 5.00
Begin VB.Form frmABCschedule 
   BackColor       =   &H00000000&
   Caption         =   "ABC Primetime Schedule"
   ClientHeight    =   6600
   ClientLeft      =   285
   ClientTop       =   1965
   ClientWidth     =   13635
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Bauhaus 93"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "VBProject.frx":0000
   ScaleHeight     =   6600
   ScaleWidth      =   13635
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   7080
      ScaleHeight     =   6195
      ScaleWidth      =   6315
      TabIndex        =   7
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton cmdGame 
      BackColor       =   &H00C0FFFF&
      Caption         =   "back to menu"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFC0&
      Caption         =   "clear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton cmdSchedule 
      BackColor       =   &H00FF8080&
      Caption         =   "display schedule"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdTime 
      BackColor       =   &H00FF8080&
      Caption         =   "find time"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton cmdDay 
      BackColor       =   &H00FF8080&
      Caption         =   "find day"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00FF8080&
      Caption         =   "Find Show"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H0080C0FF&
      Caption         =   "load schedule"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmABCschedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ctr As Integer
Dim Sname(1 To 20) As String
Dim Sday(1 To 20) As String
Dim Stime(1 To 20) As String
Dim pos As Integer

'Purpose of Form: The purpose of this form is the allow the user to view the ABC schedule.
'                 The user can also enter a favorite show, day of the week, or time and find
'                 when the shows are airing and the day.


Private Sub cmdClear_Click() 'This command button clears the picture box.
picResults.Cls
End Sub

Private Sub cmdDay_Click() 'This command button allows the user to input a day of the week.
                           'The schedule for that day is printed in the picture box.
Dim Day As String

Day = InputBox("Enter a day of the week", "Day")

For pos = 1 To Ctr
    If LCase(Day) = LCase(Sday(pos)) Then
        picResults.Print Sday(pos), Sname(pos); " airs at "; Stime(pos)
    End If
Next pos
End Sub

Private Sub cmdGame_Click() 'This command button takes the user back to the menu form.
    frmMain.Show
    frmABCschedule.Hide
End Sub

Private Sub cmdLoad_Click() 'This command button loads the ABC schedule from the data file.

cmdSchedule.Enabled = True
cmdShow.Enabled = True
cmdTime.Enabled = True
cmdDay.Enabled = True
cmdClear.Enabled = True

Open App.Path & "\ABCschedule.txt" For Input As #1

Ctr = 0
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Sname(Ctr), Sday(Ctr), Stime(Ctr)
Loop
Close #1

End Sub
Private Sub cmdSchedule_Click() 'This command button prints the entire ABC primetime schedule
                                'in the picture box.

picResults.Print "Day of Show", "Show Name"; Tab(48); "Show Time"
picResults.Print "*******************************************************************"

For pos = 1 To Ctr
    picResults.Print Sday(pos), Sname(pos); Tab(50); Stime(pos)
Next pos

End Sub

Private Sub cmdShow_Click() 'This command button asks the user for a ABC primetime show.
                            'The show's time and day is printed in the picture box.
Dim name As String
Dim found As Boolean

name = InputBox("Enter show name", "Name of show")

For pos = 1 To Ctr
    If LCase(name) = LCase(Sname(pos)) Then
        picResults.Print Sname(pos); " airs on "; Sday(pos); " at "; Stime(pos)
    End If
Next pos

End Sub

Private Sub cmdTime_Click() 'This command button asks the user for a time during primetime.
                            'The shows airing at that time and the day are printed in the picture box.

Dim Time As String

Time = InputBox("Enter primetime. Choices: 7:00, 7:30, 8:00, 8:30, or 9:00", "Time")

For pos = 1 To Ctr
    If Time = Stime(pos) Then
        picResults.Print Stime(pos), Sname(pos); " airs on "; Sday(pos)
    End If
Next pos
End Sub

 
