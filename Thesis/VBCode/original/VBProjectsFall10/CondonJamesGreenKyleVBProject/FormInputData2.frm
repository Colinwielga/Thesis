VERSION 5.00
Begin VB.Form FormInputData2 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   11295
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   FillColor       =   &H00C00000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   11295
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGradingForm 
      Caption         =   "Go To Grading Form"
      Height          =   1215
      Left            =   8880
      TabIndex        =   24
      Top             =   9360
      Width           =   2535
   End
   Begin VB.PictureBox picPic 
      Height          =   2895
      Left            =   14400
      Picture         =   "FormInputData2.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   3555
      TabIndex        =   23
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Form 2 Grade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      TabIndex        =   22
      Top             =   3000
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   5280
      ScaleHeight     =   4035
      ScaleWidth      =   12075
      TabIndex        =   21
      Top             =   4800
      Width           =   12135
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "Student Name"
      Height          =   615
      Left            =   1080
      TabIndex        =   20
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdLoops 
      Caption         =   "Use Of Loops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   19
      Top             =   9240
      Width           =   1695
   End
   Begin VB.CommandButton cmdInputBoxes 
      Caption         =   "Input from Input Boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   18
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton cmdTextBoxes 
      Caption         =   "Input From Text Boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      TabIndex        =   17
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdPictures 
      Caption         =   "Good Use of Pictues"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   16
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoodColor 
      Caption         =   "Good Use of Color"
      Height          =   855
      Left            =   12720
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdModerateColor 
      Caption         =   "Moderate Use of Color"
      Height          =   855
      Left            =   11280
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdNoColor 
      Caption         =   "No Use of Color"
      Height          =   855
      Left            =   9840
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdMultipleForms 
      Caption         =   "Multiple Forms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdArraysandSearching 
      Caption         =   "Arrays and Searching"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   10
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton cmdArraysandProcessing 
      Caption         =   "Arrays and Processing each Element in Some Way"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   9
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cmdFileOutput 
      Caption         =   "File Output"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   8
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdFileInput 
      Caption         =   "File Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   7
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdNestedStatements 
      Caption         =   "Nested If Then Statements"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoodAmount 
      Caption         =   "Good Amount of If Then Statements Used"
      Height          =   855
      Left            =   6840
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdModerate 
      Caption         =   "Moderate Amount of If Then Statements Used"
      Height          =   855
      Left            =   5160
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdzero 
      Caption         =   "No If Then Statements Used"
      Height          =   855
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox picStudentName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   2835
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblColor 
      Caption         =   "Use of Color Throughout the Project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      TabIndex        =   12
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblSelect 
      Caption         =   "Select All Features that The Student Used:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label lblIfThen 
      Caption         =   "Number of If Then Statements Used?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "FormInputData2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IfThenPoints As Integer, FeaturesPoints As Integer, ColorPoints As Integer

Private Sub cmdArraysandProcessing_Click()
'all the buttons like this when clicked give the student points for having this in their project
'It then makes it impossible for the user to click it twice
FeaturesPoints = FeaturesPoints + 5
cmdArraysandProcessing.Visible = False
End Sub

Private Sub CmdArraysandSearching_Click()
'all the buttons like this when clicked give the student points for having this in their project
'It then makes it impossible for the user to click it twice
FeaturesPoints = FeaturesPoints + 5
CmdArraysandSearching.Visible = False
End Sub

Private Sub cmdCompute_Click()
picResults.Print StudentName; " recieved a "; RunningTotal; " /50 on the first form."
picResults.Print
picResults.Print "                                                                  Points Gained From Form 2"
picResults.Print "Points Gained From If Then Statements                                      Points Gained From Color                        Points Gained From Features"
picResults.Print "**********************************************************************************************************************************************************************"
picResults.Print "     "; IfThenPoints; " / 10                                                                                          "; ColorPoints; " /10                                                      "; FeaturesPoints; " /50 "
picResults.Print
picResults.Print "**********************************************************************************************************************************************************************"
RunningTotal = RunningTotal + FeaturesPoints + IfThenPoints + ColorPoints
picResults.Print "The Total Amount of Points Earned on the Project is "; RunningTotal
cmdGradingForm.Visible = True
End Sub

Private Sub cmdFileInput_Click()
'all the buttons like this when clicked give the student points for having this in their project
'It then makes it impossible for the user to click it twice
FeaturesPoints = FeaturesPoints + 5
cmdFileInput.Visible = False
End Sub

Private Sub cmdFileOutPut_Click()
'all the buttons like this when clicked give the student points for having this in their project
'It then makes it impossible for the user to click it twice
FeaturesPoints = FeaturesPoints + 5
cmdFileOutPut.Visible = False
End Sub

Private Sub cmdGoodAmount_Click()
'all the buttons like this when clicked give the student points for having this in their project
'It then makes it impossible for the user to click it twice
IfThenPoints = IfThenPoints + 10
cmdzero.Visible = False
cmdModerate.Visible = False
End Sub

Private Sub cmdGoodColor_Click()
'all the buttons like this when clicked give the student points for having this in their project
'It then makes it impossible for the user to click it twice
ColorPoints = ColorPoints + 10
cmdModerateColor.Visible = False
cmdNoColor.Visible = False
End Sub

Private Sub cmdGradingForm_Click()
FormInputData2.Hide
FormIndividualGrades.Show

End Sub

Private Sub cmdInputBoxes_Click()
FeaturesPoints = FeaturesPoints + 5
cmdInputBoxes.Visible = False
End Sub

Private Sub cmdLoops_Click()
FeaturesPoints = FeaturesPoints + 5
cmdLoops.Visible = False
End Sub

Private Sub cmdModerate_Click()
IfThenPoints = IfThenPoints + 5
cmdzero.Visible = False
cmdGoodAmount.Visible = False

End Sub

Private Sub cmdModerateColor_Click()
ColorPoints = ColorPoints + 5
cmdGoodColor.Visible = False
cmdNoColor.Visible = False
End Sub

Private Sub cmdMultipleForms_Click()
FeaturesPoints = FeaturesPoints + 5
cmdMultipleForms.Visible = False
End Sub

Private Sub cmdName_Click()
picStudentName.Print StudentName
End Sub

Private Sub cmdNestedStatements_Click()
FeaturesPoints = FeaturesPoints + 5
cmdNestedStatements.Visible = False
End Sub

Private Sub cmdNoColor_Click()
ColorPoints = ColorPoints + 0
cmdModerateColor.Visible = False
cmdGoodColor.Visible = False
End Sub

Private Sub cmdPictures_Click()
FeaturesPoints = FeaturesPoints + 5
cmdPictures.Visible = False
End Sub

Private Sub cmdTextBoxes_Click()
FeaturesPoints = FeaturesPoints + 5
cmdTextBoxes.Visible = False
End Sub

Private Sub cmdzero_Click()
IfThenPoints = IfThenPoints + 0
cmdModerate.Visible = False
cmdGoodAmount.Visible = False
End Sub

Private Sub Form_Load()
cmdGradingForm.Visible = False
End Sub
