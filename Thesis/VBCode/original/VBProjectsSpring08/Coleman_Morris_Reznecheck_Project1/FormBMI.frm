VERSION 5.00
Begin VB.Form FormBMI 
   BackColor       =   &H00FF8080&
   Caption         =   "Form2"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form2"
   ScaleHeight     =   8295
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   720
      ScaleHeight     =   2955
      ScaleWidth      =   6315
      TabIndex        =   7
      Top             =   2160
      Width           =   6375
   End
   Begin VB.TextBox txtWeight 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtHeight 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate BMI"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back to Main Page"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   7080
      Width           =   1695
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3015
      Left            =   7440
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton cmdClickhere 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Click here to see where your BMI falls!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      MaskColor       =   &H00800000&
      TabIndex        =   0
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.Label lblBodyMass 
      BackColor       =   &H00FF8080&
      Caption         =   "Body Mass Index"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   720
      TabIndex        =   13
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label lblCalculator 
      BackColor       =   &H00FF8080&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   3120
      TabIndex        =   12
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblWeight 
      BackColor       =   &H00FF8080&
      Caption         =   "Your Weight: "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label lblLbs 
      BackColor       =   &H00FF8080&
      Caption         =   "lbs "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblHeight 
      BackColor       =   &H00FF8080&
      Caption         =   "Your Height: "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label lblCentimeters 
      BackColor       =   &H00FF8080&
      Caption         =   "inches"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   6360
      Width           =   855
   End
End
Attribute VB_Name = "FormBMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit

Private Sub cmd_Click()

End Sub

Private Sub cmdBack_Click()

frmMainpage.Show 'This hides the main page for the form
FormBMI.Hide      'BMI Calculator to show up

End Sub

Private Sub cmdCalculate_Click()

Dim ComputeBMI As Integer
Dim HeightInInches As Single
Dim WeightInPounds As Single

picResults.Cls  'Clears the screen from any previous entries

HeightInInches = txtHeight.Text
WeightInPounds = txtWeight.Text

'Formula to compute your BMI
ComputeBMI = (WeightInPounds / 2.2) / (HeightInInches / 39.4) ^ 2

    
'Once you get your result you will need to read it into an "if" statement
'There is multiple sections that tell you if your underwieght, normal or obese
'it also displays a paragraph of information that could be helpful to you.

If ComputeBMI < 19 Then
    picResults.Print "You have a BMI of " & ComputeBMI & "."
    picResults.Print
    picResults.Print " This indicates that your weight is in the underweight"
    picResults.Print " category for adults of your height. Talk with your "
    picResults.Print " healthcare provider to determine possible causes of"
    picResults.Print " underweight and if you need to gain weight."
    
    ElseIf ComputeBMI >= 19 And ComputeBMI < 25 Then
    picResults.Print "You have a BMI of " & ComputeBMI & "."
    picResults.Print
    picResults.Print " This indicates that your weight is within the normal "
    picResults.Print " range for adults of your height. Maintaining a healthy "
    picResults.Print " weight may reduce the risk of chronic diseases associated "
    picResults.Print " with overweight and obesity."
    
    ElseIf ComputeBMI >= 25 And ComputeBMI < 30 Then
    picResults.Print " You have a BMI of " & ComputeBMI & "."
    picResults.Print
    picResults.Print " This indicates that your weight is in the overweight category"
    picResults.Print " for adults of your height. People who are overweight or obese"
    picResults.Print " are at higher risk for chronic conditions such as high blood "
    picResults.Print " pressure, diabetes, and high cholesterol. Anyone who is overweight"
    picResults.Print " should try to avoid gaining additional weight. Additionally, "
    picResults.Print " if you are overweight with other risk factors (such as high "
    picResults.Print " LDL cholesterol, low HDL cholesterol, or high blood pressure), "
    picResults.Print " you should try to lose weight. Even a small weight loss (just 10% "
    picResults.Print " of your current weight) may help lower the risk of disease."
     
    ElseIf ComputeBMI >= 30 And ComputeBMI <= 40 Then
    picResults.Print "You have a BMI of " & ComputeBMI & "."
    picResults.Print
    picResults.Print " This indicates that your weight is in the obese category for"
    picResults.Print " adults of your height.People who are overweight or obese are "
    picResults.Print " at higher risk for chronic conditions such as high blood pressure,"
    picResults.Print " diabetes, and high cholesterol. At a minimum, anyone who is obese "
    picResults.Print " should try to avoid gaining additional weight. In addition, anyone"
    picResults.Print " who is obese should try to lose weight. Even a small weight loss "
    picResults.Print "(just 10% of your current weight) may help lower the risk of disease."
    picResults.Print " Talk with your healthcare provider to determine appropriate ways to "
    picResults.Print " lose weight."
    
End If

End Sub


Private Sub cmdClickhere_Click()
    picOutput.Print "BMI Range"
    picOutput.Print "___________________________________"
    picOutput.Print "Less then 19.....Underweight"
    picOutput.Print "19 to 24..........Normal/Healthy"
    picOutput.Print "25 to 29..........Overweight"
    picOutput.Print "Over 30...........Obese"
    

End Sub

Private Sub Command1_Click()
    End
End Sub


