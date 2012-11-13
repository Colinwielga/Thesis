VERSION 5.00
Begin VB.Form FrmExerciseMain 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   FillColor       =   &H0080C0FF&
   ForeColor       =   &H0080C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHeight 
      Height          =   855
      Left            =   4680
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtWeight 
      Height          =   855
      Left            =   4680
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnalyzeBMI 
      BackColor       =   &H0000FF00&
      Caption         =   "What does my BMI mean?"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdHowMany 
      BackColor       =   &H0000FF00&
      Caption         =   "How many days a week do you exercise?"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   3375
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   2280
      ScaleHeight     =   1755
      ScaleWidth      =   7515
      TabIndex        =   4
      Top             =   6240
      Width           =   7575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturntoMain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalculateBMI 
      BackColor       =   &H0000FF00&
      Caption         =   "Calculate your Body Mass Index (BMI)"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdExerciseDeifinition 
      BackColor       =   &H0000FF00&
      Caption         =   "What is it?"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label lblExercise3 
      BackColor       =   &H0080C0FF&
      Caption         =   " Exercise"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   90
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   3840
      TabIndex        =   11
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblExercise2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter your height in inches-->"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label lblExercise1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter your weight in pounds-->"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   4440
      Width           =   2655
   End
End
Attribute VB_Name = "FrmExerciseMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare global variables
Dim BMI As Single
    'Bennie Health Project
    'FrmExerciseMain
    'Heidi Donnelly
    'Written: 10/5
    'The purpose of this form is to provide the user with access to their BMI, whether it is good, the definition of exercise, types of exercise, and lastly the necessary amount of exercise needed for a healthy college-aged woman.

Private Sub cmdCalculateBMI_Click()
'this button will take the information that the user provides in the text boxes and use it to calculate their BMI

'declare variables
Dim Height As Integer
Dim Weight As Integer

'initialize variables
Height = txtHeight.Text
Weight = txtWeight.Text

'calculation
BMI = (Weight / (Height * Height)) * 703

'print results
picResults.Print "Your BMI = "; Round(BMI)
End Sub
Private Sub cmdAnalyzeBMI_Click()
'this button will ask the user for their previously calculated BMI and interpret it (say whether it's ideal, underweight, overweight, or obese)using a inputbox message and then print in the picturebox

'declare variables
Dim UserBMI As Single

'initialize variables
UserBMI = InputBox("Enter your calculated BMI")

'if/then statements determining the status of the user's BMI
If UserBMI > 30 Then
    picResults.Print "A BMI of "; Round(UserBMI); " is an indicator of obesity."
ElseIf UserBMI <= 30 And UserBMI > 25 Then
    picResults.Print "A BMI of "; Round(UserBMI); " is considered overweight."
ElseIf UserBMI <= 25 And UserBMI >= 18.5 Then
    picResults.Print "A BMI of "; Round(UserBMI); " is considered ideal."
ElseIf UserBMI < 18.5 And UserBMI > 0 Then
    picResults.Print "A BMI of "; Round(UserBMI); " is considered underweight."
Else
    picResults.Print UCase("BMI invalid!")
End If

End Sub

Private Sub cmdHowMany_Click()
'this button will use an inputbox message to ask the user how often they exercise and will then use a message box to simply explain the recommended amount of daily exercise

'declare variables
Dim Exercise As Integer

'ask user for weekly exercise
Exercise = InputBox("Enter how many days a week you exercise:")

'relay information regarding exercise through message box
MsgBox ("It is recommended that one should exercise 2-4 days a week for sessions of about 30 minutes.")
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturntoMain_Click()
    MsgBox ("I hope you got the results you were looking for.")
    FrmExerciseMain.Hide
    FrmMain.Show
End Sub
Private Sub cmdExerciseDeifinition_Click()
'this button provides the user with the definition of the word exercise
    MsgBox ("Exercise = is any movement that works your body at a greater intensity than your usual level of daily activity. Exercise raises your heart rate and works your muscles and is most commonly undertaken to achieve the aim of physical fitness.") & Chr(13) & ("(Examples of exercise include: walking, jogging, biking, swimming, dancing, yoga, tennis, boxing, gardening, and workout DVDs.)")
End Sub


