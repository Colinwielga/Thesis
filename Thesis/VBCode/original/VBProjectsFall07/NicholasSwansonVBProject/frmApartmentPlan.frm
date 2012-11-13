VERSION 5.00
Begin VB.Form frmApartmentPlan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Apartment Plan"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Results"
      Height          =   855
      Left            =   9240
      TabIndex        =   14
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Retrieve Weekly Punch and Bucks Totals Based on Your Spending Habits"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5400
      TabIndex        =   13
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton cmdDescription 
      Caption         =   "Description"
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtPunchInput 
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtBucksInput 
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate Projections"
      Height          =   855
      Left            =   1080
      MaskColor       =   &H80000005&
      TabIndex        =   4
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   4440
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
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   120
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
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   5400
      ScaleHeight     =   915
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   3120
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
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblPunchUse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please type expected number of punches to be used each week:"
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
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   2775
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
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblStepOne 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmApartmentPlan.frx":0000
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
      Left            =   5640
      TabIndex        =   8
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label lblFoodFight 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Food Fight 2007: Apartment Plan"
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
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmApartmentPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form calculates the week and month when the Apartment Plan will be exhausted of
'punches and dining bucks, according to the user's expected weekly spending habits.
'It does so by calculating the number of weeks it takes for punches and bucks to max
'out and comparing this number to an array of data received from one of two text files,
'one containing the weeks and months of the fall semester, the other containing the weeks
'and months of the spring semester.  It also allows the user to check weekly remaining
'totals of punches and bucks based upon their expected weekly spending.  This is accomplished
'through comparison of the week being searched for with the calculations made according to
'the user's expected weekly spending habits.

Option Explicit
'Option Explicit to avoid variable confusion.  Define global variables
Dim Ctr As Integer, WeeksMonths(1 To 100) As String, MonthCounterPunches As Integer, MonthCounterBucks As Integer, PunchNumber As Double, BucksNumber As Double
Private Sub chkFall_Click()
'Ensures that only one checkbox can be selected at a time
If chkFall.Value = 1 Then
    chkSpring.Value = 0     'deselects chkSpring when chkFall is selected
End If
End Sub

Private Sub chkSpring_Click()
'Ensures that only one checkbox can be selected at a time
If chkSpring.Value = 1 Then
    chkFall.Value = 0       'deselects chkFall when ChkSpring is selected
End If
End Sub

Private Sub cmdBack_Click()
'Returns user to Start Page
frmApartmentPlan.Hide 'closes frmApartmentPlan
frmStartPage.Show     'opens frmStartPage
End Sub
Private Sub cmdCalculate_Click()
'Dim and identify the variables
Dim PunchTotal As Single, BucksTotal As Single
'Ensure that user cannot cause error by not entering values in the textboxes
If Val(txtPunchInput.Text) <= 0 Or Val(txtBucksInput.Text) <= 0 Then    'If inputted values are <= 0, error message will display
    MsgBox "Please enter values in both input boxes", , "Need a value"
Else:                                                           'If the inputted values are > 0, the program will do the computations as described
    'Define inputs and corresponding variables
    PunchNumber = txtPunchInput.Text
    BucksNumber = txtBucksInput.Text
    'Load Arrays
    Ctr = 0
    If chkFall.Value = 1 Then                                   'chkFall loads file of fall weeks and months, to be searched later
        Open App.Path & "\FallSchedule.txt" For Input As #1     'open text file FallSchedule.txt
        Do While Not EOF(1)                                     'Load file into array WeeksMonths until all items in file are loaded
            Ctr = Ctr + 1
            Input #1, WeeksMonths(Ctr)
        Loop                                                    'Repeat until all data is loaded
        Close #1                                                'Close SpringSchedule.txt when done load data into array WeeksMonths
    Else: Open App.Path & "\SpringSchedule.txt" For Input As #1 'chkSpring loads file of spring weeks and months, to be searched later; opens text file SpringSchedule.txt
        Do While Not EOF(1)                                     'Load file into array WeeksMonths until all items in file are loaded
            Ctr = Ctr + 1
            Input #1, WeeksMonths(Ctr)
        Loop                                                    'Repeat until all data is loaded
        Close #1                                                'Close SpringSchedule.txt when done load data into array WeeksMonths
    End If                                                      'End the If statement
    'Compute Results
    PunchTotal = 0
    BucksTotal = 0
    MonthCounterPunches = 0
    MonthCounterBucks = 0
    Do While PunchTotal < 70
        MonthCounterPunches = MonthCounterPunches + 1           'MonthCounterPunches keeps track of how many weeks will pass before punches are exhausted
        PunchTotal = PunchNumber * MonthCounterPunches          'PunchTotal keeps track of how many punches are spent over time
    Loop                                                        'When PunchTotal is >= 150, the loop stops and MonthCounterPunches is stored for later use
    Do While BucksTotal < 200
        MonthCounterBucks = MonthCounterBucks + 1               'MonthCounterBucks keeps track of how many weeks will pass before dining bucks are exhausted
        BucksTotal = BucksNumber * MonthCounterBucks            'BucksTotal keeps track of how many bucks are spent over time
    Loop                                                        'When BucksTotal is >= 175, the loops stops and MonthCounterBucks is stored for later use
    'Clear previous results
    picResults.Cls
    'Print Results
    If MonthCounterPunches > 17 Then
        picResults.Print "You will not run out of punches before the end of the semester."                  'If MonthCounterPunches > 17, punches will not run out before the semester ends
        ElseIf PunchNumber > 70 Then                                                                       'Prints error statement if PunchNumber exceeds plan's allotted number of punches
            picResults.Print "Error: entered value exceeds allotted number of punches."
        Else: picResults.Print "Your punches will last until the "; WeeksMonths(MonthCounterPunches); "."   'If MonthCounterPunches <= 17, PunchTotal must have been >= 150 before the end of the semester,
    End If                                                                                                  'so the final week in WeeksMonths to be reached according to MonthCounterPunches is printed as the result.
    If MonthCounterBucks > 17 Then
        picResults.Print "You will not run out of dining bucks before the end of the semester."              'If MonthCounterBucks > 17, dining bucks will not run out before the semester ends
        ElseIf BucksNumber > 200 Then                                                                        'Prints error statement if BucksNumber exceeds plan's allotted number of dining bucks
            picResults.Print "Error: entered value exceeds allotted number of dining bucks."
        Else: picResults.Print "Your dining bucks will last until the "; WeeksMonths(MonthCounterBucks); "." 'If MonthCounterBucks <= 17, BucksTotal must have been >= 175 before the end of the semester,
    End If                                                                                                   'so the final week in WeeksMonths to be reached according to MonthCounterBucks is printed as the result.
End If
'activates cmdSearch, now that the arrays have been loaded and MonthCounterPunches/MonthCounterBucks calculated
    cmdSearch.Enabled = True
End Sub

Private Sub cmdClear_Click()
'clears picResults
picResults.Cls
End Sub

Private Sub cmdDescription_Click()
'Gives Details about the Block Plan
MsgBox "Available only to Students living in Apartment style living.  Provides 70 meals per semester versus meals per week.  This plan includes 200 Dining Bucks for the semester.", , "Apartment Plan Description"
End Sub

'Searches WeeksMonths array according to MonthCounterPunches and MonthCounterBucks
'to return the remaining number of punches and bucks for any week within these specified
'numbers of weeks.

Private Sub cmdSearch_Click()
'dim variables
Dim WeekOfMonth As String, FoundPunches As Boolean, FoundBucks As Boolean, WeekCtrPunches As Integer, WeekCtrBucks As Integer
Dim PunchRemainder As Integer, BucksRemainder As Integer

'define inputbox, in which is typed a week and month in the same style as FallSchedule.txt and SpringSchedule.txt
WeekOfMonth = InputBox("Please enter the week and month for which you would like your remaining punch and flex totals.  Use the format 'first week of October'.", "Weekly Totals Search")

WeekCtrPunches = 0      'set new counters equal to 0
WeekCtrBucks = 0
FoundPunches = False    'set Boolean variables equal to False
FoundBucks = False
Do While ((Not FoundPunches) And (WeekCtrPunches < MonthCounterPunches))    'Search for a specific week within MonthCounterPunches.  Loop the search until the week is found or WeekCtrPunches reaches MonthCounterPunches.
    WeekCtrPunches = WeekCtrPunches + 1
    If WeekOfMonth = WeeksMonths(WeekCtrPunches) Then FoundPunches = True   'If the entered string = a week and month string in WeeksMonths as defined by MonthCounterPunches, values can be found for the entered string
Loop
Do While ((Not FoundBucks) And (WeekCtrBucks < MonthCounterBucks))          'Search for a specific week within MonthCounterBucks.  Loop the search until the week is found or WeekCtrBucks reaches MonthCounterBucks.
    WeekCtrBucks = WeekCtrBucks + 1
    If WeekOfMonth = WeeksMonths(WeekCtrBucks) Then FoundBucks = True       'If the entered string = a week and month string in WeeksMonths as defined by MonthCounterBucks, values can be found for the entered string
Loop

PunchRemainder = 70 - PunchNumber * WeekCtrPunches
BucksRemainder = 200 - BucksNumber * WeekCtrBucks

If Not FoundPunches Then                                                                'If the entered string does not exist within WeeksMonths as defined by MonthCounterPunches, a messagebox will say so
    picResults.Print "You will have no remaining punches at this time."
Else: picResults.Print PunchRemainder; " punches will remain to you at this time."      'If the entered string does exit within this range, caculate the number of punches remaining and print this value
End If
If Not FoundBucks Then                                                                      'If the entered string does not exist within WeeksMonths as defined by MonthCounterBucks, a messagebox will say so
    picResults.Print "You will have no remaining dining bucks at this time."
Else: picResults.Print BucksRemainder; " dining bucks will remain to you at this time."     'If the entered string does exit within this range, caculate the number of dining bucks remaining and print this value
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
