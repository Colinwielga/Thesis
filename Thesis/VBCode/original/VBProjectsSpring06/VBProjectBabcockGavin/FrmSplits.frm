VERSION 5.00
Begin VB.Form frmSplits 
   BackColor       =   &H00400000&
   Caption         =   "Splits Page"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   6960
      TabIndex        =   10
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdSplitByName 
      BackColor       =   &H00FF8080&
      Caption         =   "Calculate Runner's 5k splits"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   8235
      TabIndex        =   8
      Top             =   2280
      Width           =   8295
   End
   Begin VB.CommandButton cmdInfoByName 
      BackColor       =   &H00FF8080&
      Caption         =   "Search by Runner's Name"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdNavigate1 
      Caption         =   "Go To Results Page"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdSplits 
      Caption         =   "Calculate Split Times"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtSeconds 
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtMinutes 
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblAuthors 
      BackColor       =   &H00400000&
      Caption         =   " By Steven Babcock and Sam Gavin                                                         23 March, 2006"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   5400
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   5040
      Picture         =   "FrmSplits.frx":0000
      Top             =   240
      Width           =   3315
   End
   Begin VB.Label lblOr 
      Caption         =   "OR"
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Minutes:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblSeconds 
      Caption         =   "Seconds:"
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmSplits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim globally to be used within multiple subcommands in this form only
Dim N As String
Dim MinuteLapSplit As Single
Dim SecondLapSplit As Single
Dim LapSplit As Single

Private Sub cmdInfoByName_Click(Index As Integer)
    Dim Found As Boolean
    Dim Pos As Integer
    N = InputBox("Enter the name of the desired runner you wish to find", "Name")
    Pos = 0
    Found = False
    
    Do While ((Not Found) And (Pos < ArraySize))
        Pos = Pos + 1
        If N = Names(Pos) Then Found = True
    Loop
    
    If (Not Found) Then
            picResults.Print N; Names(Pos); " did not run this event"
        Else
            picResults.Print N; " finished in position "; Place(Pos); "with a time of "; Minutes(Pos); ":"; Seconds(Pos)
    End If
End Sub


Private Sub cmdNavigate1_Click()
    'Go to results page
    frmResults.Show
    'Hide splits page
    frmSplits.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSplitByName_Click()
    Dim Found As Boolean
    Dim Pos As Integer
    picResults.Cls
    'Have user enter runner name
    N = InputBox("Enter the name of the desired runner you wish to find", "Name")
    Pos = 0
    'Set initial boolean command found to false
    Found = False
    'Use boolean to search until the name is found
    Do While ((Not Found) And (Pos < ArraySize))
        Pos = Pos + 1
        If N = Names(Pos) Then Found = True
    Loop
    
    If (Not Found) Then
            picResults.Print N; Names(Pos); " did not run this event"
        Else
            'Convert minutes into seconds then divide by the 12.5 laps that a 5k is run in on a normal 400m track
            MinuteLapSplit = (Minutes(Pos) * 60) / 12.5
            'now divide seconds
            SecondLapSplit = Seconds(Pos) / 12.5
            LapSplit = MinuteLapSplit + SecondLapSplit
            'Lap splits are commonly read in seconds in the running world even when larger than one minute.
            picResults.Print N; " Average Lap Split was "; LapSplit; " Seconds "
    End If
End Sub

Private Sub cmdSplits_Click()
    Dim SplitMinutes As Single
    Dim SplitSeconds As Single
    
    'Assign variables to the two text boxes
    SplitMinutes = txtMinutes.Text
    SplitSeconds = txtSeconds.Text
    
    'Clear the picture box
    picResults.Cls
    
    'Convert minutes into seconds then divide by the 12.5 laps that a 5k is run in on a normal 400m track
    MinuteLapSplit = (SplitMinutes * 60) / 12.5
    'now divide seconds
    SecondLapSplit = SplitSeconds / 12.5
    LapSplit = MinuteLapSplit + SecondLapSplit
    'Lap splits are commonly read in seconds in the running world even when larger than one minute.
    picResults.Print N; " Average Lap Split was "; LapSplit; " Seconds "

End Sub


