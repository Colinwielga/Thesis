VERSION 5.00
Begin VB.Form frmCalculate 
   BackColor       =   &H00800000&
   Caption         =   "Select a Measure"
   ClientHeight    =   7395
   ClientLeft      =   2220
   ClientTop       =   2205
   ClientWidth     =   10530
   FillColor       =   &H80000007&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10530
   Begin VB.CommandButton cmdFlex 
      Caption         =   "Evaluate my Flexibility"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   14
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdLowerBody 
      Caption         =   "Evaluate my Lower Body Strength"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   13
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdTargetHR 
      Caption         =   "Calculate my Target Heart Rate"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   11
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdUpperBody 
      Caption         =   "Evaluate my Upper Body Strength"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   9
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdBMR 
      Caption         =   "Calculate my BMR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
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
      Left            =   1320
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   6000
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
      Height          =   975
      Left            =   7080
      TabIndex        =   4
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdBMI 
      Caption         =   "Calculate my BMI"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
      Caption         =   "Flexibilty"
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
      Height          =   495
      Left            =   5880
      TabIndex        =   15
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
      Caption         =   "Lower Body Strength"
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
      Height          =   495
      Left            =   5880
      TabIndex        =   12
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      Caption         =   "Target Heart Rate"
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
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "Upper Body Strength"
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
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "BMR (Basal Metabolic Rate)"
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
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "BMI (Body Mass Index)"
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
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Select a Measure to Calculate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Click on one of the buttons below to calculate the indicated measure of your health and fitness."
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
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
   End
End
Attribute VB_Name = "frmCalculate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmCalculate
'Nick Schuster
'March 26, 2008

'This form lets the user select which measure he/she would
'like to calculate and then directs the user to the appropriate form.
Option Explicit
'To return to the previous form
Private Sub cmdBack_Click()
frmCalculate.Hide
frmInfo.Show
End Sub
'Makes sure the necessary information has been entered for the selected measure
Private Sub cmdBMI_Click()
If Inches = 0 Or Weight = 0 Then
    frmCalculate.Hide
    frmInfo.Show        'Returns the user to the information form and tells the user which information is needed
    MsgBox "You must fill in the Height and Weight boxes to calculate your BMI.", , "Fill in all necessary boxes"
Else
    frmCalculate.Hide   'Takes the user to the next form, if necessary information is in place
    frmBMI.Show
End If


End Sub
'Makes sure the necessary information has been entered for the selected measure
Private Sub cmdBMR_Click()
If Inches = 0 Or Weight = 0 Or Age = 0 Or (Gender <> "M" And Gender <> "F") Then
    frmCalculate.Hide
    frmInfo.Show        'Returns the user to the information form and tells the user which information is needed
    MsgBox "You must fill in the Gender, Age, Height and Weight, boxes to calculate your BMR. Gender must be M or F (boxes are case sensitive).", , "Fill in all necessary boxes"
Else
    frmCalculate.Hide   'Takes the user to the next form, if necessary information is in place
    frmBMR.Show
End If

End Sub
'Makes sure the necessary information has been entered for the selected measure
Private Sub cmdFlex_Click()
If Age = 0 Or (Gender <> "M" And Gender <> "F") Then
    frmCalculate.Hide
    frmInfo.Show        'Returns the user to the information form and tells the user which information is needed
    MsgBox "You must fill in the Gender, Age, and Trunk Flexibility boxes to evaluate your flexibility. Gender must be M or F (boxes are case sensitive).", , "Fill in all necessary boxes"
Else
    frmCalculate.Hide   'Takes the user to the next form, if necessary information is in place
    frmFlexibility.Show
End If

End Sub
'Makes sure the necessary information has been entered for the selected measure
Private Sub cmdLowerBody_Click()
If (Gender <> "M" And Gender <> "F") Or Age = 0 Or Weight = 0 Or MaxLeg = 0 Then
    frmCalculate.Hide
    frmInfo.Show        'Returns the user to the information form and tells the user which information is needed
    MsgBox "You must fill in the Gender, Age, Weight and Maximum Leg Press boxes to evaluate your lower body strength. Gender must be M or F (boxes are case sensitive).", , "Fill in all necessary boxes"
Else
    frmCalculate.Hide   'Takes the user to the next form, if necessary information is in place
    frmLowerBodyStrength.Show
End If
End Sub
'To quit the program now
Private Sub cmdQuit_Click()
End
End Sub
'Makes sure the necessary information has been entered for the selected measure
Private Sub cmdTargetHR_Click()

If Age = 0 Or RestingHeartRate = 0 Then
    frmCalculate.Hide
    frmInfo.Show        'Returns the user to the information form and tells the user which information is needed
    MsgBox "You must fill in the Age and Resting Heart Rate boxes to calculate your target heart rate.", , "Fill in all necessary boxes"
Else
    frmCalculate.Hide   'Takes the user to the next form, if necessary information is in place
    frmTargetHR.Show
End If


End Sub
'Makes sure the necessary information has been entered for the selected measure
Private Sub cmdUpperBody_Click()
If (Gender <> "M" And Gender <> "F") Or Age = 0 Or Weight = 0 Or MaxBench = 0 Then
    frmCalculate.Hide
    frmInfo.Show        'Returns the user to the information form and tells the user which information is needed
    MsgBox "You must fill in the Gender, Age, Weight and Maximum Bench Press boxes to evaluate your upper body strength. Gender must be M or F (boxes are case sensitive).", , "Fill in all necessary boxes"
Else
    frmCalculate.Hide   'Takes the user to the next form, if necessary information is in place
    frmUpperBodyStrength.Show
End If
End Sub
