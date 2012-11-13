VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00800000&
   Caption         =   "Basic Information"
   ClientHeight    =   9075
   ClientLeft      =   2430
   ClientTop       =   1425
   ClientWidth     =   10380
   LinkTopic       =   "Form2"
   ScaleHeight     =   9075
   ScaleWidth      =   10380
   Begin VB.CommandButton cmdWhatFlex 
      Caption         =   "What's this?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   23
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtFlexibility 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7440
      TabIndex        =   22
      Text            =   "0"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtMaxLeg 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7440
      TabIndex        =   21
      Text            =   "0"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtMaxBench 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7440
      TabIndex        =   20
      Text            =   "0"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtRestingHeartRate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7440
      TabIndex        =   19
      Text            =   "0"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtWeight 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3120
      TabIndex        =   18
      Text            =   "0"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtHeight 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3120
      TabIndex        =   17
      Text            =   "0"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3120
      TabIndex        =   16
      Text            =   "0"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtGender 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3120
      TabIndex        =   15
      Text            =   "0"
      Top             =   3000
      Width           =   1335
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
      Left            =   4560
      TabIndex        =   13
      Top             =   7800
      Width           =   1455
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
      Left            =   1080
      MaskColor       =   &H8000000F&
      TabIndex        =   12
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
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
      MaskColor       =   &H8000000F&
      TabIndex        =   11
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton cmdWhatLeg 
      Caption         =   "What's this?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdWhatBench 
      Caption         =   "What's this?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdWhatHR 
      Caption         =   "What's this?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00800000&
      Caption         =   "**You do not need to fill in all the boxes in order to proceed, but if you do not know what to fill in, PLEASE LEAVE IT AS ""0"".**"
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
      Height          =   615
      Left            =   1920
      TabIndex        =   25
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00800000&
      Caption         =   "Your trunk flexibility score in inches:"
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
      Height          =   615
      Left            =   4920
      TabIndex        =   24
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00800000&
      Caption         =   "Your gender (Enter M or F):"
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
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Your age in years:"
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
      TabIndex        =   7
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
      Caption         =   "Your height in inches:"
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
      TabIndex        =   6
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
      Caption         =   "Your maximum bench press in pounds:"
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
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      Caption         =   "Your maximum leg press in pounds:"
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
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "Your weight in pounds:"
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
      TabIndex        =   3
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Your resting heart rate in beats per minute:"
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
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   $"frmInfo.frx":0000
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
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Basic Information"
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
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmInfo
'Nick Schuster
'March 26, 2008

'This form gathers basic physical information from the user about him/herself and enters that data into the computer.
Option Explicit
'To go back to the last form
Private Sub cmdBack_Click()
frmInfo.Hide
frmWelcome.Show
End Sub

Private Sub cmdContinue_Click()
'To assign the values entered by the user to variables which will be used for making calculations later on
Gender = txtGender.Text
Age = Round(txtAge.Text)
Inches = Round(txtHeight.Text)
Weight = Round(txtWeight.Text)
RestingHeartRate = Round(txtRestingHeartRate.Text)
MaxBench = Round(txtMaxBench.Text)
MaxLeg = Round(txtMaxLeg.Text)
TrunkFlex = Round(txtFlexibility.Text)
'To go on to the next form
frmInfo.Hide
frmCalculate.Show
End Sub
'To quit the program now
Private Sub cmdQuit_Click()
End
End Sub
'To inquire about the indicated measure
Private Sub cmdWhatBench_Click()
frmInfo.Hide
frmexplain2.Show
End Sub
'To inquire about the indicated measure
Private Sub cmdWhatFlex_Click()
frmInfo.Hide
frmExplain4.Show
End Sub
'To inquire about the indicated measure
Private Sub cmdWhatHR_Click()
frmInfo.Hide
frmExplain1.Show
End Sub
'To inquire about the indicated measure
Private Sub cmdWhatLeg_Click()
frmInfo.Hide
frmExplain3.Show
End Sub


