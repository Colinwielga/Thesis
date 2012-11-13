VERSION 5.00
Begin VB.Form frmLowerBodyStrength 
   BackColor       =   &H00800000&
   Caption         =   "Your Lower Body Strength Evaluation"
   ClientHeight    =   7815
   ClientLeft      =   3210
   ClientTop       =   1440
   ClientWidth     =   8640
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   8640
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Evaluate My Lower Body Strength"
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
      Top             =   6360
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
      Top             =   6360
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
      Top             =   6360
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
      Left            =   1680
      ScaleHeight     =   1095
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   1920
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Click the button  below to Evaluate your Lower Body Strength"
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
      Top             =   960
      Width           =   6735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Maximum Leg Press"
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
      Height          =   975
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
      Caption         =   $"frmLowerBodyStrength.frx":0000
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
      Height          =   2895
      Left            =   1680
      TabIndex        =   4
      Top             =   3000
      Width           =   5175
   End
End
Attribute VB_Name = "frmLowerBodyStrength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmLowerBodyStrength
'Nick Schuster
'March 26, 2008

'This form compares the user's lower body strength to that of the general population of his/her age and gender
Option Explicit
'To return to the previous form
Private Sub cmdBack_Click()
frmLowerBodyStrength.Hide
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
WeightRatio = MaxLeg / Weight   'Calculates the user's strength to weight ratio

picResults.Cls
picResults.Print "You can leg press "; FormatPercent(WeightRatio, 0); " of your body weight." 'Reports ratio as a percentage

If Gender = "M" Then                                        'This section of code determines which data file to open based
                                                            'on the user's gender (with an If-Then statement) and age (using
    Select Case Age                                         'a Select Case statement). Once the appropriate data file has been
        Case Is < 20                                        'determined, the program opens it.
            Open App.Path & "\LPM10s.txt" For Input As #1
        Case 20 To 29
            Open App.Path & "\LPM20s.txt" For Input As #1
        Case 30 To 39
            Open App.Path & "\LPM30s.txt" For Input As #1
        Case 40 To 49
            Open App.Path & "\LPM40s.txt" For Input As #1
        Case 50 To 59
            Open App.Path & "\LPM50s.txt" For Input As #1
        Case Is >= 60
            Open App.Path & "\LPM60s.txt" For Input As #1
    End Select
        
ElseIf Gender = "F" Then
    
    Select Case Age
        Case Is < 20
            Open App.Path & "\LPW10s.txt" For Input As #1
        Case 20 To 29
            Open App.Path & "\LPW20s.txt" For Input As #1
        Case 30 To 39
            Open App.Path & "\LPW30s.txt" For Input As #1
        Case 40 To 49
            Open App.Path & "\LPW40s.txt" For Input As #1
        Case 50 To 59
            Open App.Path & "\LPW50s.txt" For Input As #1
        Case Is >= 60
            Open App.Path & "\LPW60s.txt" For Input As #1
    End Select

End If

Do While Not EOF(1)                     'Reads the correct data file into two arrays using a Do While loop
    Ctr = Ctr + 1
    Input #1, Rating(Ctr), Score(Ctr)
Loop

Close #1

For J = Ctr To 1 Step -1                'Determined which rating the user should receive by comparing the user's
    If WeightRatio >= Score(J) Then     'weight ratio to the lower bound of each rating in the data file
        YourRating = Rating(J)
    End If
Next J

picResults.Print "This indicates "; YourRating; " lower body strength."
End Sub
'To quit the program now
Private Sub cmdQuit_Click()
End
End Sub

