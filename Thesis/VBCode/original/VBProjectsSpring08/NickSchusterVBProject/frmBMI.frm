VERSION 5.00
Begin VB.Form frmBMI 
   BackColor       =   &H00800000&
   Caption         =   "Your BMI"
   ClientHeight    =   8055
   ClientLeft      =   2385
   ClientTop       =   1605
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   8865
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate BMI"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   3
      Top             =   6720
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   1920
      ScaleHeight     =   975
      ScaleWidth      =   5415
      TabIndex        =   2
      Top             =   1680
      Width           =   5415
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
      Left            =   600
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   6720
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
      Left            =   6120
      TabIndex        =   0
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Click the button  below to Calculate your BMI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Body Mass Index (BMI)"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   $"frmBMI.frx":0000
      ForeColor       =   &H8000000E&
      Height          =   3855
      Left            =   1920
      TabIndex        =   4
      Top             =   2760
      Width           =   5415
   End
End
Attribute VB_Name = "frmBMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmBMI
'Nick Schuster
'March 26, 2008

'This form calculates the user's BMI
Option Explicit
'To return to the previous form
Private Sub cmdBack_Click()
frmBMI.Hide
frmCalculate.Show
End Sub

Private Sub cmdCalculate_Click()
Dim BMI As Single
Dim J As Single
Dim WeightMin As Single
Dim WeightMax As Single

picResults.Cls

BMI = (Weight / 2.2) / ((Inches * 2.54) / 100) ^ 2          'Calculates the users BMI via the standard equation
picResults.Print "Your BMI is "; FormatNumber(BMI, 1); "."
    Select Case BMI
        Case Is < 18.5                                                             'This section of code uses a Select Case
            picResults.Print "This indicates that you are underweight."            'statement to determine the user's weight
        Case 18.5 To 24.9                                                          'category and report it.
            picResults.Print "This indicates that you are at a healthy weight."
        Case 25 To 29
            picResults.Print "This indicates that you are overweight."
        Case 30 To 34.9
            picResults.Print "This indicates that you have grade 1 obesity (significantly overweight)."
        Case 35 To 39.9
            picResults.Print "This indicates that you have grade 2 obesity (severely overweight)."
        Case Is > 40
            picResults.Print "This indicates that you have grade 3 obesity (morbidly/massively overweight)."
        Case Else
            frmBMI.Hide
            frmInfo.Show
            MsgBox "Please recheck the information you entered for your height and weight.", , "Invalid Information Entry"
    End Select

End Sub
'To quit the program now
Private Sub cmdQuit_Click()
End
End Sub



