VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00400000&
   Caption         =   "Statistics Page"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   8520
      TabIndex        =   4
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdNavigate1 
      BackColor       =   &H00FF8080&
      Caption         =   "Go To Results Page"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdRange 
      BackColor       =   &H00FF8080&
      Caption         =   "Find the range of all results"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdAverage 
      BackColor       =   &H00FF8080&
      Caption         =   "Find average of all results"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdMedian 
      BackColor       =   &H00FF8080&
      Caption         =   "Find median of all results"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblNames 
      BackColor       =   &H00400000&
      Caption         =   "By Sam Gavin and Steven Babcock                                                                         23 March, 2006"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   8040
      Width           =   9135
   End
   Begin VB.Image ImgSpeedySteve 
      Height          =   5355
      Left            =   720
      Picture         =   "FrmStats.frx":0000
      Top             =   1800
      Width           =   9000
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MinutesTotal As Single
Dim SecondsTotal As Single
Dim TimeCounter As Integer
Dim Pos As Integer
Dim Median As Integer
Private Sub cmdAverage_Click()
Dim AvgMinutes As Integer
Dim AvgSeconds As Single
    
    'set pos to zero
    Pos = 0
    'clear output variables so they work after clicking more than once
    AvgMinutes = 0
    AvgMinutes = 0
    'Open the file
        Open App.Path & "\RunnerResults.txt" For Input As #1

    'begin loop to load into the array and print each line in the text file
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Place(Pos), Names(Pos), Year(Pos), School(Pos), Minutes(Pos), Seconds(Pos)
    Loop
    
    TimeCounter = 0
    For Pos = 1 To ArraySize
        MinutesTotal = MinutesTotal + Minutes(Pos)
        SecondsTotal = SecondsTotal + Seconds(Pos)
    Next Pos
    AvgMinutes = MinutesTotal / ArraySize
    AvgSeconds = SecondsTotal / ArraySize
    MsgBox "The average time was " & AvgMinutes & ":" & FormatNumber(AvgSeconds), , "Average Runner Time"
    'Close the file
    Close #1

End Sub

Private Sub cmdMedian_Click()
    'comupte the median using a numeric function
    Median = ArraySize / 2
    MsgBox "The Median Time is" & Minutes(Median) & ":" & Seconds(Median), , "Median Results"
End Sub

Private Sub cmdNavigate1_Click()
    frmResults.Show
    frmStats.Hide

End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRange_Click()
    'Reload the text file and re-write the array so it's organized by time
    'set counter (pos) at zero
    Pos = 0
    'Open the file
    Open App.Path & "\RunnerResults.txt" For Input As #1
    
    'begin loop to load into the array and print each line in the text file
    Do While Not EOF(1)
        Pos = Pos + 1
        Input #1, Place(Pos), Names(Pos), Year(Pos), School(Pos), Minutes(Pos), Seconds(Pos)
    Loop
    'Close the file
    Close #1
    
    'Since the text file is already organized by time (you obviously record results as the runners pass the finish line)
    'We can simply print the first and last times in the text file
    MsgBox "The Range is: " & Minutes(1) & ":" & Seconds(1) & " minutes to " & Minutes(ArraySize) & ":" & Seconds(ArraySize) & " Minutes"

End Sub

Private Sub cmdNavigate3_Click()
    frmResults.Show
    frmStats.Hide
    
End Sub
