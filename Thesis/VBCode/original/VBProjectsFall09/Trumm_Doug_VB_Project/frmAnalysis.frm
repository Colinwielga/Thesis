VERSION 5.00
Begin VB.Form frmAnalysis 
   BackColor       =   &H00004040&
   Caption         =   "Dealing with the workers"
   ClientHeight    =   10920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14175
   FillColor       =   &H00004080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Moving on Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12360
      TabIndex        =   7
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10440
      TabIndex        =   6
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search by Age"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7320
      TabIndex        =   5
      Top             =   9000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   6855
      Left            =   8520
      Picture         =   "frmAnalysis.frx":0000
      ScaleHeight     =   6795
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Alphabetize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3960
      TabIndex        =   3
      Top             =   9000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      Height          =   6975
      Left            =   480
      ScaleHeight     =   6915
      ScaleWidth      =   6555
      TabIndex        =   2
      Top             =   1560
      Width           =   6615
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Worker Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   1
      Top             =   9000
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004040&
      Caption         =   "the office"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   13815
   End
End
Attribute VB_Name = "frmAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'This program will open a file, put it into an array, and sort it and perform other operations on the data
    'These include alphabetizing and match/stop searching
    
    'Declare universal varibales for form
    Dim workers(1 To 25) As String, ages(1 To 25) As Integer, wCTR As Integer



Private Sub cmdOpen_Click()
    'Declare variables
    
    'Prepare the file to be opened
    Open App.Path & "\workers.txt" For Input As #1
    wCTR = 0
    picResults.Print "Worker"; Tab(20); "Age"
    picResults.Print "**************************************************************************************************"
    
    'Open file with a Do While Loop
    Do While Not EOF(1)
        wCTR = wCTR + 1
        Input #1, workers(wCTR), ages(wCTR)
        picResults.Print workers(wCTR); Tab(20); ages(wCTR)
    Loop
    
    Close #1
    
    'Make other buttons visibles after data has been sorted into arrays
    cmdAlpha.Visible = True
    cmdOpen.Visible = False
    cmdSearch.Visible = True
    
End Sub

Private Sub cmdAlpha_Click()

    'Declare variables
    Dim Pass As Integer, Pos As Integer, TempW As String, TempA As Integer, J As Integer

    'Use bubble sort fuction to alphabetize the list of names.  Temporary values are used to facilitate the swap.
    For Pass = 1 To wCTR - 1
        For Pos = 1 To wCTR - Pass
            If workers(Pos) > workers(Pos + 1) Then
                TempW = workers(Pos)
                TempA = ages(Pos)
                workers(Pos) = workers(Pos + 1)
                ages(Pos) = ages(Pos + 1)
                workers(Pos + 1) = TempW
                ages(Pos + 1) = TempA
            End If
        Next Pos
    Next Pass
    
    'Print results neatly
    picResults.Cls
    picResults.Print "Worker"; Tab(20); "Age"
    picResults.Print "**************************************************************************************************"
    For J = 1 To wCTR
        picResults.Print workers(J); Tab(20); ages(J)
    Next J
    
End Sub

Private Sub cmdSearch_Click()
    'Declare variables
    Dim ageInput As Single, J As Integer, Found As Boolean
    
    'Input box allow user to select the age to look for
    ageInput = InputBox("Enter an age to search for amongst the list of workers")
    Found = False
    J = 0
    
    'Use Match/Stop in a Do While Loop to search only until a match is found
    Do While Not Found And J < wCTR
        J = J + 1
        If ageInput = ages(J) Then
            Found = True
        End If
    Loop
    
    'If statements indicate which result occurred
    If Found = True Then
        MsgBox (workers(J) & " is " & ageInput & " years old.")
    Else
        MsgBox ("There are no workers that are " & ageInput & " years old in the list.")
    End If
    
    cmdSwitch.Visible = True
End Sub

Private Sub cmdStop_Click()
    End
End Sub

Private Sub cmdSwitch_Click()
    'Move on to next form
    frmAnalysis.Hide
    frmEvasion.Show
End Sub
