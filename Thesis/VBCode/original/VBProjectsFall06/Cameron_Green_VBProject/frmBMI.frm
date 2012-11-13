VERSION 5.00
Begin VB.Form frmBMI 
   BackColor       =   &H00008000&
   Caption         =   "Body Mass Index Test"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      ScaleHeight     =   1875
      ScaleWidth      =   9915
      TabIndex        =   2
      Top             =   3960
      Width           =   9975
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H000000FF&
      Caption         =   "Calculate Your Body Mass Index"
      Height          =   1815
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C00000&
      Caption         =   "Back to Homepage"
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3240
      Left            =   3840
      Picture         =   "frmBMI.frx":0000
      Top             =   360
      Width           =   3240
   End
End
Attribute VB_Name = "frmBMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to homepage from BMI page'
Private Sub cmdBack_Click()
    frmIntroCC.Show
    frmBMI.Hide
End Sub

Private Sub cmdCalculate_Click()
picResults.Cls
'clear picResults picture box'
Dim BMI As Single, Weight As Integer, Height As Integer
Height = InputBox("Enter your Height in Inches")
Weight = InputBox("Enter your Weight in Pounds")
'user enters in height and weight using input boxes'
BMI = (Weight / (Height * Height)) * 703
'this step calculates the body mass index using the formula above'
Select Case BMI
    Case Is < 18.5
    picResults.Print "Your BMI of"; BMI; "is considered to be underweight."
    Case Is >= 18.5 And BMI < 25
    picResults.Print "Your BMI of"; BMI; "is considered a normal weight."
    Case Is >= 25 And BMI < 30
    picResults.Print "Your BMI of"; BMI; "is considered to be overweight."
    Case Else
    picResults.Print "Your BMI of"; BMI; "is considered to be obese."
End Select
'this case select statement tells the program that if the BMI is in a certain range, then it will print out a certain response for each BMI'
End Sub
