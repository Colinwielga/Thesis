VERSION 5.00
Begin VB.Form frmFlexibility 
   BackColor       =   &H00800000&
   Caption         =   "Your Flexibility Evaluation"
   ClientHeight    =   7125
   ClientLeft      =   2715
   ClientTop       =   1440
   ClientWidth     =   8385
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   8385
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
      Height          =   975
      Left            =   1560
      ScaleHeight     =   975
      ScaleWidth      =   5415
      TabIndex        =   3
      Top             =   1800
      Width           =   5415
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
      Height          =   975
      Left            =   5880
      TabIndex        =   2
      Top             =   5760
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
      Height          =   975
      Left            =   360
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Evaluate My Flexibility"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Click the button  below to Evaluate your Flexibility"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Trunk Flexibility Score"
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
      Height          =   1215
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
      Caption         =   $"frmMaxHR.frx":0000
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
      Height          =   2535
      Left            =   1560
      TabIndex        =   4
      Top             =   2880
      Width           =   5175
   End
End
Attribute VB_Name = "frmFlexibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmFlexibility
'Nick Schuster
'March 26, 2008

'This form compares the user's flexibility score to that of the general population of his/her age and gender
Option Explicit
'To go back to the previous form
Private Sub cmdBack_Click()
frmFlexibility.Hide
frmCalculate.Show
End Sub

Private Sub cmdCalculate_Click()
Dim Score(1 To 10) As Single     'Dims variables that will refer
Dim Rating(1 To 10) As String    'to arrays, counters, and output
Dim Ctr As Integer
Dim J As Integer
Dim YourRating As String


Ctr = 0

picResults.Cls
picResults.Print "Your trunk flexibility score is "; TrunkFlex; " inches."

If Gender = "M" Then
    
    Select Case Age                                            'This portion of the code determines the appropriate data
        Case Is < 26                                           'file to open based on the user's gender (with an If-Then
            Open App.Path & "\TFM0025.txt" For Input As #1     'statement) and age (with a Select Case statement). Once
        Case 26 To 35                                          'the appropriate data file has been determined, the program
            Open App.Path & "\TFM2635.txt" For Input As #1     'opens it.
        Case 36 To 45
            Open App.Path & "\TFM3645.txt" For Input As #1
        Case 46 To 55
            Open App.Path & "\TFM4655.txt" For Input As #1
        Case 56 To 65
            Open App.Path & "\TFM5665.txt" For Input As #1
        Case Is > 65
            Open App.Path & "\TFM6600.txt" For Input As #1
    End Select
        
ElseIf Gender = "F" Then
    
    Select Case Age
        Case Is < 26
            Open App.Path & "\TFW0025.txt" For Input As #1
        Case 26 To 35
            Open App.Path & "\TFW2635.txt" For Input As #1
        Case 36 To 45
            Open App.Path & "\TFW3645.txt" For Input As #1
        Case 46 To 55
            Open App.Path & "\TFW4655.txt" For Input As #1
        Case 56 To 65
            Open App.Path & "\TFW5665.txt" For Input As #1
        Case Is > 65
            Open App.Path & "\TFW6600.txt" For Input As #1
    End Select

End If

Do While Not EOF(1)                     'The program reads the correct data file into two arrays
    Ctr = Ctr + 1                       'with a Do While Loop
    Input #1, Rating(Ctr), Score(Ctr)
Loop

Close #1

For J = Ctr To 1 Step -1                'The program determines which range of possible scores the
    If TrunkFlex >= Score(J) Then       'user fits into by comparing the user's score with the lower
        YourRating = Rating(J)          'bound of each rating listed in the data file
    End If
Next J

picResults.Print "This indicates "; YourRating; " flexibility." 'The program reports the user's rating

End Sub
'To quit the program now
Private Sub cmdQuit_Click()
End
End Sub

