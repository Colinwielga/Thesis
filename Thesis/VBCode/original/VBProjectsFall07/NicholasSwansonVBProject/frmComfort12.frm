VERSION 5.00
Begin VB.Form frmComfort12 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Comfort 12 Plan"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Results"
      Height          =   855
      Left            =   9120
      TabIndex        =   13
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Retrieve Weekly Punch and Bucks Totals Based on Your Spending Habits"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5280
      TabIndex        =   12
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CommandButton cmdDescription 
      Caption         =   "Description"
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtBucksInput 
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate Projections"
      Height          =   855
      Left            =   960
      MaskColor       =   &H80000005&
      TabIndex        =   4
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CheckBox chkFall 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fall"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   8.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox chkSpring 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Spring"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   5280
      ScaleHeight     =   915
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label lblChooseSemester 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please choose a semester:"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblPunchUse 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "You will receive 12 punches each week and 110 dining bucks for the semester."
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label lblBucksUse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please type expected number of dining bucks to be spent each week:"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblStepOne 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmComfort12.frx":0000
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1455
      Left            =   5520
      TabIndex        =   7
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label lblFoodFight 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Food Fight 2007: Comfort 12 Plan"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmComfort12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form calculates the week and month when the Comfort 12 Plan will be
'exhausted of dining bucks, according to the user's expected weekly spending habits.
'It does so by calculating the number of weeks it takes for bucks to max out and
'comparing this number to an array of data received from one of two text files,
'one containing the weeks and months of the fall semester, the other containing the weeks
'and months of the spring semester.  It also allows the user to check weekly remaining
'totals of bucks based upon their expected weekly spending.  This is accomplished
'through comparison of the week being searched for with the calculations made according to
'the user's expected weekly spending habits.

Option Explicit
'Option Explicit to avoid variable confusion.  Dim the universal variables
Dim Ctr As Integer, WeeksMonths(1 To 100) As String, MonthCounterBucks As Integer, BucksNumber As Double

Private Sub chkFall_Click()
'Deactivates the Spring checkbox when the Fall checkbox is active
If chkFall.Value = 1 Then
    chkSpring.Value = 0     'deselects chkSpring when chkFall is selected
End If
End Sub

Private Sub chkSpring_Click()
'Deactivates Fall checkbox when Spring checkbox is active
If chkSpring.Value = 1 Then
    chkFall.Value = 0       'deselects chkFall when chkSpring is selected
End If
End Sub

Private Sub cmdBack_Click()
'returns user to frmChooseComfortType
frmComfort12.Hide           'closes frmComfort12
frmChooseComfortType.Show   'opens frmChooseComfortType
End Sub

Private Sub cmdCalculate_Click()
'Dim and identify the variables
Dim BucksTotal As Single
'Ensure that user cannot cause error by not entering value in the textbox
If Val(txtBucksInput.Text) <= 0 Then                                    'If inputted values are <= 0, error message will display
    MsgBox "Please enter a value in the input box", , "Need a value"
Else:                                                           'If the inputted values are > 0, the program will do the computations as described
    'Define input and corresponding variable
    BucksNumber = txtBucksInput.Text
    'Load Arrays for number of weeks in each month
    Ctr = 0
    If chkFall.Value = 1 Then
        Open App.Path & "\FallSchedule.txt" For Input As #1     'open the read file FallSchedule.txt if Fall checkbox is active
        Do While Not EOF(1)                                     'execute loop until the final data entry in the file
            Ctr = Ctr + 1                                       'advance counter (Ctr) by 1 each time the loop is repeated
            Input #1, WeeksMonths(Ctr)                          'Load the read file into the array WeeksMonths according to Ctr
        Loop                                                    'repeat if Do While Not conditions are not met
        Close #1                                                'close the read file
    Else: Open App.Path & "\SpringSchedule.txt" For Input As #1 'open the read file Spring Schedule.txt if Sprint checkbox is active
        Do While Not EOF(1)                                     'execute loop until the final data entry in the file
            Ctr = Ctr + 1                                       'advance the counter (Ctr) by 1 each time the loop is repeated
            Input #1, WeeksMonths(Ctr)                          'Load the read file into the array WeeksMonths according to Ctr
        Loop                                                    'repeat if Do While Not conditions are not met
        Close #1                                                'close the read file
    End If
    'Compute Results
    BucksTotal = 0
    MonthCounterBucks = 0
    Do While BucksTotal < 110                                  'Define Do While condition
        MonthCounterBucks = MonthCounterBucks + 1               'advance MonthCounterBucks by 1 each time the loop is repeated
        BucksTotal = BucksNumber * MonthCounterBucks            'multiply BucksNumber by MonthCounterBucks to determine when Do While condition is met
    Loop                                                        'repeat loop if Do While condition is not met
    'Print Results
    'Clear Previous Results
    picResults.Cls
    'New punches are issued every week, so no calculations are necessary
    picResults.Print "You will receive 12 new punches each week."
    'If MonthCounterBucks exceeds 17 weeks, dining bucks will not run out before the end of the semester
    If MonthCounterBucks > 17 Then
        picResults.Print "You will not run out of dining bucks before the end of the semester."
        ElseIf BucksNumber > 110 Then                                                                           'Prints error statement if BucksNumber exceeds plan's allotted number of dining bucks
                    picResults.Print "Error: entered value exceeds allotted number of dining bucks."
        Else: picResults.Print "Your dining bucks will last until the "; WeeksMonths(MonthCounterBucks); "."    'If MonthCounterBucks is 17 or less, print the value from WeeksMonths that corresponds to the integer value of MonthCounterBucks
    End If
End If
'activates cmdSearch, now that the arrays have been loaded and MonthCounterPunches/MonthCounterBucks calculated
    cmdSearch.Enabled = True
End Sub

Private Sub cmdClear_Click()
'clears picResults
picResults.Cls
End Sub

Private Sub cmdDescription_Click()
'Gives Details about the Comfort 12 Plan
MsgBox "12 Meals per week & 110 Dining Bucks per semester", , "Comfort 12 Plan Description"
End Sub

'Searches WeeksMonths array according to MonthCounterBucks to return
'the remaining number of punches and bucks for any week within these specified
'numbers of weeks.

Private Sub cmdSearch_Click()
'dim variables
Dim WeekOfMonth As String, FoundBucks As Boolean, WeekCtrBucks As Integer, BucksRemainder As Integer

'define inputbox, in which is typed a week and month in the same style as FallSchedule.txt and SpringSchedule.txt
WeekOfMonth = InputBox("Please enter the week and month for which you would like your remaining punch and flex totals.  Use the format 'first week of October'.", "Weekly Totals Search")

WeekCtrBucks = 0        'set new counters equal to 0
FoundBucks = False      'set Boolean variables equal to False
Do While ((Not FoundBucks) And (WeekCtrBucks < MonthCounterBucks))          'Search for a specific week within MonthCounterBucks.  Loop the search until the week is found or WeekCtrBucks reaches MonthCounterBucks.
    WeekCtrBucks = WeekCtrBucks + 1
    If WeekOfMonth = WeeksMonths(WeekCtrBucks) Then FoundBucks = True       'If the entered string = a week and month string in WeeksMonths as defined by MonthCounterBucks, values can be found for the entered string
Loop
If Not FoundBucks Then                                                                      'If the entered string does not exist within WeeksMonths as defined by MonthCounterBucks, a messagebox will say so
    MsgBox "You will have no remaining dining bucks at this time.", , "Meal Plan Exhausted"
ElseIf BucksRemainder > 110 Then
    MsgBox "You will have no remaining dining bucks at this time.", , "Meal Plan Exhausted"
Else: BucksRemainder = 110 - BucksNumber * WeekCtrBucks                                     'If the entered string does exit within this range, caculate the number of dining bucks remaining and print this value
    picResults.Print BucksRemainder; " dining bucks will remain to you at this time."
End If
End Sub

Private Sub Form_Load()
'According to cmdCalculate, if neither box is active then txtSpringSchedule will load.
'This subroutine prevents accidental loading of the wrong file by automatically activating one checkbox,
'which ensures that the user will purposefully pick one or the other.
chkFall.Value = 1
'centers form on computer screen upon loading
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
