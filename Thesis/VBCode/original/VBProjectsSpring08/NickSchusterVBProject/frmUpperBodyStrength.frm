VERSION 5.00
Begin VB.Form frmUpperBodyStrength 
   BackColor       =   &H00800000&
   Caption         =   "Your Upper Body Strength Evaluation"
   ClientHeight    =   7560
   ClientLeft      =   3210
   ClientTop       =   1605
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   8610
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Evaluate My Upper Body Strength"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   3
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6000
      TabIndex        =   1
      Top             =   6120
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   1800
      ScaleHeight     =   1095
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Click the button  below to Evaluate your Upper Body Strength"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Maximum Bench Press"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
      Caption         =   $"frmUpperBodyStrength.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3135
      Left            =   1800
      TabIndex        =   4
      Top             =   2880
      Width           =   5175
   End
End
Attribute VB_Name = "frmUpperBodyStrength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmUpperBodyStrength
'Nick Schuster
'March 26, 2008

'This form compares the user's lower body strength to that of the general population of his/her age and gender
Option Explicit
'To return to the previous form
Private Sub cmdBack_Click()
frmUpperBodyStrength.Hide
frmCalculate.Show
End Sub

Private Sub cmdCalculate_Click()

Dim Score(1 To 10) As Single     'Dims variables that will refer to arrays, counters, and output
Dim Rating(1 To 10) As String
Dim WeightRatio As Single
Dim Ctr As Integer
Dim J As Integer
Dim YourRating As String


Ctr = 0
WeightRatio = MaxBench / Weight     'Calculates the user's strength to weight ratio

picResults.Cls
picResults.Print "You can bench press "; FormatPercent(WeightRatio, 0); " of your body weight."     'Displays the ration as a percentage

If Gender = "M" Then
    
    Select Case Age                                         'Determines which data file to open based on the user's gender
        Case Is < 20                                        '(using an If-Then Statement) adn age (using a Select Case
            Open App.Path & "\BPM10s.txt" For Input As #1   'statement). Once the appropriate file has been determined, the
        Case 20 To 29                                       'program opens it.
            Open App.Path & "\BPM20s.txt" For Input As #1
        Case 30 To 39
            Open App.Path & "\BPM30s.txt" For Input As #1
        Case 40 To 49
            Open App.Path & "\BPM40s.txt" For Input As #1
        Case 50 To 59
            Open App.Path & "\BPM50s.txt" For Input As #1
        Case Is >= 60
            Open App.Path & "\BPM60s.txt" For Input As #1
    End Select
        
ElseIf Gender = "F" Then
    
    Select Case Age
        Case Is < 20
            Open App.Path & "\BPW10s.txt" For Input As #1
        Case 20 To 29
            Open App.Path & "\BPW20s.txt" For Input As #1
        Case 30 To 39
            Open App.Path & "\BPW30s.txt" For Input As #1
        Case 40 To 49
            Open App.Path & "\BPW40s.txt" For Input As #1
        Case 50 To 59
            Open App.Path & "\BPW50s.txt" For Input As #1
        Case Is >= 60
            Open App.Path & "\BPW60s.txt" For Input As #1
    End Select

End If

Do While Not EOF(1)                     'Uses a Do While Loop to load the correct data file into two arrays
    Ctr = Ctr + 1
    Input #1, Rating(Ctr), Score(Ctr)
Loop

Close #1

For J = Ctr To 1 Step -1                'Uses a For-Next loop to determine the user's rating by comparing the
    If WeightRatio >= Score(J) Then     'user's ratio to the lower bound of each rating in the data file
        YourRating = Rating(J)
    End If
Next J

picResults.Print "This indicates "; YourRating; " upper body strength."

End Sub
'To quit the program now
Private Sub cmdQuit_Click()
End
End Sub

