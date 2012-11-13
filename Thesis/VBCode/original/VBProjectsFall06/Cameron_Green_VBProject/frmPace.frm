VERSION 5.00
Begin VB.Form frmPace 
   BackColor       =   &H00008000&
   Caption         =   "Pace Calculator"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Go Back to Homepage"
      Height          =   1095
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   480
      ScaleHeight     =   5595
      ScaleWidth      =   9315
      TabIndex        =   2
      Top             =   1800
      Width           =   9375
   End
   Begin VB.CommandButton cmdCalories 
      BackColor       =   &H000000FF&
      Caption         =   "Calculate how Many Calories you Burned"
      Height          =   1095
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdCalculateMiles 
      BackColor       =   &H000000FF&
      Caption         =   "Calculate your pace per mile"
      Height          =   1095
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmPace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to homepage from pace calculator page'
Private Sub cmdBack_Click()
    frmIntroCC.Show
    frmPace.Hide
End Sub

Private Sub cmdCalculateMiles_Click()
picResults.Cls
'clears the picResults screen every time you want to calculate your pace'
Dim Miles As Single, Time As Single, Total As Single
Miles = InputBox("Enter how many miles you ran.")
Time = InputBox("Enter your time for your run.")
'user enters in miles ran and time it took into input boxes'
Total = Time / Miles
'the pace per mile is calculated using the formula above'
picResults.Print "You ran"; Total; "minutes per mile"
'the pace per mile is printed out in the picResults picture box'
End Sub

Private Sub cmdCalories_Click()
picResults.Cls
'clears the picResults screen every time you want to calculate how many calories you burned'
Dim Miles As Single, Weight As Integer, Calories As Double
Miles = InputBox("Enter how many miles you ran.")
Weight = InputBox("Enter your weight in pounds.")
'user enters miles ran and weight in pounds via input boxes'
Calories = (Miles * Weight * 1.036)
'this calculates how many calories are burned using the formula above'
picResults.Print "Your burned"; Calories; "calories today."
'this prints out how many calories were burned in the picResults picture box'
End Sub
