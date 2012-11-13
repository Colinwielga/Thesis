VERSION 5.00
Begin VB.Form trackcalculations 
   BackColor       =   &H00800000&
   Caption         =   "Calculate data essential to have a successful track workout program"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdestimate 
      BackColor       =   &H0000C000&
      Caption         =   "Estimate your maximum effort for a race you don't have an official time for "
      Height          =   1815
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CommandButton cmdtarget 
      BackColor       =   &H000000FF&
      Caption         =   "Calculate your target heart rate"
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton cmdpace 
      BackColor       =   &H00FF0000&
      Caption         =   "Calculate your workout pace"
      Height          =   1695
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton cmdswitch 
      BackColor       =   &H0000FFFF&
      Caption         =   "Compare Your Time To The Best"
      Height          =   1215
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   3015
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      Height          =   1215
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   2655
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H80000009&
      Height          =   5295
      Left            =   3720
      ScaleHeight     =   5235
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Track and Field Workout Calculations"
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   9735
   End
End
Attribute VB_Name = "trackcalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : trackconversions (Van Heel, Tom.vbp)
'Form Name : trackcalculations (Van Heel, Tom.frm)
'Author: Tom Van Heel
'Date Written: March 12, 2004
'Purpose of Form: Do do basic calculations necessary to run a successful
                 'sprinter workout.  Calculations include target heart rate,
                 'which is the number of heartbeats per minute which are necessary
                 'to have for a workout to have a positive cardiovascular impact.
                 'The workout percentage calculator is necessary for determining
                 'what pace a runner should be at given they want to run a certain
                 'percent of their workout.  Finally, the coversion formula makes
                 'it possible to determine what your estimated maximum effort would
                 'be for a race that you don't have an official time for, which is
                 'great for overdistance training as well as gaining footspeed with
                 'shorter workouts.

Private Sub cmdestimate_Click()
'This button is discribed above, and uses a formula that takes the ratio of
'previously coverted times, and multiplies it with the time entered by the user.
Dim K As Single
Dim T As Single
Dim E As Single
Dim PATH As String
Dim distance(1 To 6) As Single
Dim seconds(1 To 6) As Single
Dim original As Single
Dim convert As Single
Dim newtime As Single

K = Val(InputBox("Enter the distance of your race", "100, 200, or 400"))
T = Val(InputBox("Enter your best time in seconds", "race time"))
E = Val(InputBox("Enter the distance you wish to convert your time to", "100, 150, 200, 300, 400, 500"))
'input boxes allow the user to enter data, without taking up space on the form screen
'as with a text box
PATH = "N:\CS130\handin\Van Heel, Tom\"
picresults.Cls 'clears the picture box screen, which makes use of the program a lot more user friendly.

'The following string of code opens a file and searches through it to find data that
'matches what the user entered into the input boxes.  Then it sets the found variables
'with a new name so it can find it later.
Open PATH & "conversions.txt" For Input As #3
    For D = 1 To 6
    Input #3, distance(D), seconds(D)
        If distance(D) = K Then
            original = seconds(D)
        ElseIf distance(D) = E Then
            convert = seconds(D)
        End If
    Next D
    
'The following formula takes variables that were previously determined and
'calculates the new time from it.
If original <> 0 Then
    newtime = (convert / original) * T
End If

picresults.Print "Your estimated time is "; FormatNumber(newtime, 2) 'prints results
Close #3 'closes file

End Sub

Private Sub cmdpace_Click()
'the following code takes the time entered by the user, and divides it by the
'percent that the user entered to determine the pace that the user should run in.
Dim time As Double
Dim percent As Double
Dim workout As Double
T = InputBox("Enter your best race time in seconds", "Best race") 'user enters data
    If T < 0 Then
        MsgBox ("Time must be greater than zero.")
        'an error message that prompts the user to correct their entry
        T = InputBox("Enter your best race time in seconds", "Best race")
    End If
percent = InputBox("Enter the percent you intend to run your workout in as a decimal.", "Percent between 0.00 and 1.00")
picresults.Cls
    If percent > 1 Or percent < 0 Then 'This is used to make sure the the number entered is a percentage
        MsgBox "You cannot enter a percentage greater than one or less than zero", , "Error"
        percent = InputBox("Enter the percent you intend to run your workout in as a decimal", "Percent")
    End If
workout = T / percent 'this calculates the workout time
picresults.Print "Your workout pace should be "; FormatNumber(workout, 2); " seconds."
End Sub

Private Sub cmdquit_Click()
'allows the user to quit anytime
End
End Sub

Private Sub cmdswitch_Click()
'used to switch between the two different forms
trackcalculations.Hide
trackcomparison.Show
End Sub

Private Sub cmdtarget_Click()
'This string of code is used to calculate the target heartrate of the user
'based off the age they enter.
Dim age As Single
Dim lower As Single
Dim upper As Single
Dim max As Single
picresults.Cls
age = InputBox("Enter your age.", "Age")
    If age < 0 Then 'does not allow the user to enter a negative number as their age
        MsgBox ("Age must be greater than zero")
        age = InputBox("Enter your age.", "Age")
    End If
'The following calculation is not the perfect measure of what a person's target
'heart rate should be, but it is a fairly good estimate.  To accurately determine
'target heart rate, consult your physician.
max = 220 - age
lower = max * 0.65
upper = max * 0.85
picresults.Print "Your target heartrate is between "; lower; " and "; upper; " beats per minute."

End Sub


