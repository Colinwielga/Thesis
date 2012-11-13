VERSION 5.00
Begin VB.Form FormGroupInputData2 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   12645
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12645
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForm3 
      Caption         =   "Go To Grading Form"
      Height          =   735
      Left            =   7440
      TabIndex        =   28
      Top             =   11760
      Width           =   3135
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Form 2 Totals"
      Height          =   1335
      Left            =   16200
      TabIndex        =   27
      Top             =   5040
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   3495
      Left            =   720
      ScaleHeight     =   3435
      ScaleWidth      =   15795
      TabIndex        =   26
      Top             =   7920
      Width           =   15855
   End
   Begin VB.CommandButton cmdGoodPic 
      Caption         =   "Extensive Use of Pictures/Images"
      Height          =   735
      Left            =   15120
      TabIndex        =   25
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSomePic 
      Caption         =   "Some Use Of Pictures/Images"
      Height          =   855
      Left            =   13440
      TabIndex        =   24
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdNoPic 
      Caption         =   "No Use Of Pictures/Images"
      Height          =   855
      Left            =   11880
      TabIndex        =   23
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdGoodColor 
      Caption         =   "Extensive Use Of Color"
      Height          =   735
      Left            =   9840
      TabIndex        =   21
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdSomeColor 
      Caption         =   "Some Use Of Color"
      Height          =   735
      Left            =   8640
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdNoColor 
      Caption         =   "No Use Of Color"
      Height          =   735
      Left            =   7320
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdString 
      Caption         =   "Use of String Functions"
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
      Left            =   10440
      TabIndex        =   17
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdMessageBox 
      Caption         =   "Use of Message Boxes"
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
      Left            =   8160
      TabIndex        =   16
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdMath 
      Caption         =   "Use of Math Funtions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   15
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdFileOutPut 
      Caption         =   "Use of File Output"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdFileInput 
      Caption         =   "Use of File Input"
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
      TabIndex        =   13
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdArraysAndSorting 
      Caption         =   "Use of Arrays and Sorting"
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
      Left            =   12960
      TabIndex        =   12
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton CmdArraysandSearching 
      Caption         =   "Use Of Arrays And Searching"
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
      Left            =   5760
      TabIndex        =   11
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoops 
      Caption         =   "Use Of Loops "
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
      Left            =   3240
      TabIndex        =   10
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdModules 
      Caption         =   "Use Of Module Variables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12960
      TabIndex        =   9
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdNestedIfs 
      Caption         =   "Use Of Nested Ifs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10560
      TabIndex        =   8
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdInputBoxes 
      Caption         =   "Input From Input Boxes"
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
      Left            =   8040
      TabIndex        =   7
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdTextBoxes 
      Caption         =   "Input From Text Boxes"
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
      Left            =   600
      TabIndex        =   6
      Top             =   5040
      Width           =   2055
   End
   Begin VB.PictureBox picName2 
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
      Left            =   3120
      ScaleHeight     =   915
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.PictureBox picName1 
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
      Left            =   480
      ScaleHeight     =   915
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdStudentNames 
      Caption         =   "Students' Names"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblPic 
      Caption         =   "Use of Pictures/Images Throughout Project"
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
      Left            =   12240
      TabIndex        =   22
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblColor 
      Caption         =   "Use of Color Throughout Project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   18
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblSelect 
      Caption         =   "Select all Features Included:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Label lblStudent2 
      Caption         =   "Student 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblName1 
      Caption         =   "Student 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "FormGroupInputData2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColorPoints As Integer, PicPoints As Integer, FeaturesPoints As Integer

Private Sub CmdArraysandSearching_Click()
CmdArraysandSearching.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdArraysAndSorting_Click()
cmdArraysAndSorting.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdCompute_Click()
'this prints out the points earned from the different imput from the user and adds up the running total
picResults.Print StudentName; " and "; StudentName2; " recieved a "; RunningTotal; " out of 70 on the first form"
picResults.Print
picResults.Print "                                                                          Points Earned From Form 2"
picResults.Print "Points Earned From Color                                Points Earned From Pictures/Images                     Points Earned From included Featues"
picResults.Print ColorPoints; " out of 10                                                           "; PicPoints; " out of 10                                                         "; FeaturesPoints; " Points Earned "
picResults.Print "*********************************************************************************************************************************************************************************"
RunningTotal = RunningTotal + FeaturesPoints + ColorPoints + PicPoints
picResults.Print StudentName; " and "; StudentName2; " recieved a total of "; RunningTotal; "for the entire project"
cmdForm3.Visible = True
End Sub

Private Sub cmdFileInput_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdFileInput.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdFileOutPut_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdFileOutPut.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdForm3_Click()
FormGroupInputData2.Hide
FormGroupGrades.Show
End Sub

Private Sub cmdGoodColor_Click()
'assings points for the color used in the project and makes it impossible to select a different buttong once this has been selected
cmdNoColor.Visible = False
cmdSomeColor.Visible = False
ColorPoints = 10
End Sub

Private Sub cmdGoodPic_Click()
'assings points for the pictures used in the project and makes it impossible to select a different buttong once this has been selected
cmdNoPic.Visible = False
cmdSomePic.Visible = False
PicPoints = 10
End Sub

Private Sub cmdInputBoxes_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdInputBoxes.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdLoops_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdLoops.Visible = False
FeaturesPoints = FeaturesPoints + 5

End Sub

Private Sub cmdMath_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdMath.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdMessageBox_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdMessageBox.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdModules_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdModules.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdNestedIfs_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdNestedIfs.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdNoColor_Click()
'assings points for the color used in the project and makes it impossible to select a different buttong once this has been selected
cmdSomeColor.Visible = False
cmdGoodColor.Visible = False
ColorPoints = 0
End Sub

Private Sub cmdNoPic_Click()
'assings points for the pictures used in the project and makes it impossible to select a different buttong once this has been selected
cmdSomePic.Visible = False
cmdGoodPic.Visible = False
PicPoints = 0

End Sub

Private Sub cmdSomeColor_Click()
'assings points for the color used in the project and makes it impossible to select a different buttong once this has been selected
cmdNoColor.Visible = False
cmdGoodColor.Visible = False
ColorPoints = 5
End Sub

Private Sub cmdSomePic_Click()
'assings points for the pictures used in the project and makes it impossible to select a different buttong once this has been selected
cmdNoPic.Visible = False
cmdGoodPic.Visible = False
PicPoints = 5
End Sub

Private Sub cmdString_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdString.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub cmdStudentNames_Click()
picName1.Print StudentName
picName2.Print StudentName2
End Sub

Private Sub cmdTextBoxes_Click()
'gives the student points when the user selects this and adds it to the Features points
cmdTextBoxes.Visible = False
FeaturesPoints = FeaturesPoints + 5
End Sub

Private Sub Form_Load()
cmdForm3.Visible = False
End Sub
