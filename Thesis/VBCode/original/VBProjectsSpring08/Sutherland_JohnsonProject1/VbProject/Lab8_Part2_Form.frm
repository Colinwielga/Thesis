VERSION 5.00
Begin VB.Form BMI_Index 
   BackColor       =   &H00000000&
   Caption         =   "BMI index"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack5 
      Caption         =   "Back to Main Page"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdTotalBMI 
      BackColor       =   &H00C00000&
      Caption         =   "What's you BMI?"
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   6675
      TabIndex        =   4
      Top             =   840
      Width           =   6735
   End
   Begin VB.TextBox txtWeight 
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtFeetTall 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblWeight 
      Caption         =   "How much do you weigh in pounds?"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblFeetTall 
      Caption         =   "How tall are you in inches?"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "BMI_Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BMI Index Code
'Calorie Counter
'By: Ryan Sutherland and Tara Johnson
'2-26-08
'This form is used to calculate a person Body Mass Index based upon your height and weight.
Private Sub cmdBack5_Click()
Main_Page.Show
BMI_Index.Hide
MsgBox "Thank you for taking the time and effort to make sure that you are healthy. The programmers hope that you are happy with the program and will continue to use it to figure out your calories. Thank you again.", , "Thank you"
End Sub
'Ends the program
Private Sub cmdQuit_Click()
End
End Sub
'This button take a height in inches input by the user and weight in pounds and calculates if
'they are healthy according to the Body Mass Index.
Private Sub cmdTotalBMI_Click()
Dim FeetTall As Single
Dim Weight As Single
FeetTall = txtFeetTall.Text                   'Text box to enter you height in inches for the below calculation.
Weight = txtWeight.Text                       'Text box to enter you weight in pounds for the below calculation.
TotalBMI = (Weight / 2.2) / (FeetTall / 39.4) ^ 2
'This is an If Statement.  It will first find out if the variable in question is or is not going to fit the
'requirement that is listed.  If it does, it will do the steps listed below it.
If TotalBMI < 19 Then
    picResults.Print "Your BMI is"; FormatNumber(TotalBMI, 3); "which is considered unsafe and may indicate malnourishment"
ElseIf TotalBMI >= 19 And TotalBMI < 25 Then
    picResults.Print "Your BMI is"; FormatNumber(TotalBMI, 3); "which is considered a healthy weight."
ElseIf TotalBMI >= 25 And TotalBMI < 30 Then
    picResults.Print "Your BMI is"; FormatNumber(TotalBMI, 3); "which is considered overweight."
ElseIf TotalBMI >= 30 And TotalBMI < 40 Then
    picResults.Print "Your BMI is"; FormatNumber(TotalBMI, 3); "which is considered obese."
ElseIf TotalBMI > 40 Then
    picResults.Print "Your BMI is"; FormatNumber(TotalBMI, 3); "which is considered very obese."
End If
'The formatNumber() in the above if statements is used to make sure that the number that is typed in will only go
'to three decimal places.
End Sub

