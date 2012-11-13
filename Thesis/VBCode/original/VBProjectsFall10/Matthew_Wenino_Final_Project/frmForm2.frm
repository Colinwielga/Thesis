VERSION 5.00
Begin VB.Form frmForm2 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   9870
   ClientLeft      =   3780
   ClientTop       =   2550
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   7695
   Begin VB.CommandButton cmdNextPage 
      Caption         =   "Next Page"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      TabIndex        =   11
      Top             =   7920
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5400
      TabIndex        =   10
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox txtHeightInch 
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtWeight 
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtHeightFt 
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      ScaleHeight     =   2115
      ScaleWidth      =   6675
      TabIndex        =   4
      Top             =   3480
      Width           =   6735
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4440
      TabIndex        =   1
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdHomePage 
      Caption         =   "Home Page"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3480
      TabIndex        =   0
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label lblWeight 
      Caption         =   "Please Enter Your Weight Here"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label lblHeight 
      Caption         =   "Please Enter Your Height Here (feet, Inch)"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lblIntro 
      Caption         =   $"frmForm2.frx":0000
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()
Dim BMI As Single, Weight As Long, Height As Integer, HeightFt As Integer, HeightInch As Integer
Weight = txtWeight.Text
HeightFt = txtHeightFt.Text
HeightInch = txtHeightInch.Text

Height = (HeightFt * 12) + HeightInch
BMI = (Weight * 703) / Height ^ 2
If BMI < 18.5 Then
    picResults.Print "Your BMI is "; FormatNumber(BMI, 2); "."
ElseIf BMI > 18.5 And BMI < 24.9 Then
    picResults.Print "Your BMI is "; FormatNumber(BMI, 2); "."
ElseIf BMI > 25 And BMI < 29.9 Then
    picResults.Print "Your BMI is "; FormatNumber(BMI, 2); "."
ElseIf BMI >= 30 Then
    picResults.Print "Your BMI is "; FormatNumber(BMI, 2); "."
End If
End Sub

Private Sub cmdClear_Click()
picResults.Cls
txtWeight.Text = " "
txtHeightFt.Text = " "
txtHeightInch.Text = " "
End Sub

Private Sub cmdHomePage_Click()
frmForm1.Show
frmForm2.Hide
End Sub

Private Sub cmdNextPage_Click()
frmForm2.Hide
frmForm3.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
