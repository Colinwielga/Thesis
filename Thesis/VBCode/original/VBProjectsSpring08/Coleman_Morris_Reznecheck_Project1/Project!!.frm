VERSION 5.00
Begin VB.Form frmTargetHeartRate 
   BackColor       =   &H00FF8080&
   Caption         =   "Target Heart Rate Calculator"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      ScaleHeight     =   315
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00E0E0E0&
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
      Height          =   1095
      Left            =   480
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate Your Target Heart Rate"
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
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   3600
      Width           =   4215
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtAge 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtResting 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00FF8080&
      Caption         =   "and"
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
      Left            =   6240
      TabIndex        =   18
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblYourHR 
      BackColor       =   &H00FF8080&
      Caption         =   "Your Heart rate is  Between "
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
      Left            =   4920
      TabIndex        =   17
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "4. The usual resting rate for an adult is between 50 to 100 beats per minute."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   7200
      Width           =   7455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "3. Count the number of beats for 30 seconds; then double the result to get the number of beats per minute."
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
      Left            =   240
      TabIndex        =   13
      Top             =   6480
      Width           =   7335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "2. Place two fingers (not your thumb) gently over your carotid artery"
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
      Left            =   240
      TabIndex        =   12
      Top             =   6120
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "1. Rest quietly for 10 minutes."
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
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "How to Calculate Your Resting Heart Rate:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   4815
   End
   Begin VB.Label lblBeats 
      BackColor       =   &H00FF8080&
      Caption         =   "beats per minute"
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
      Left            =   5040
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblYears 
      BackColor       =   &H00FF8080&
      Caption         =   "yrs"
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
      Left            =   5040
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblAge 
      BackColor       =   &H00FF8080&
      Caption         =   "Age: "
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
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblRestingHeartRate 
      BackColor       =   &H00FF8080&
      Caption         =   "Resting Heart Rate: "
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   2655
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
      Left            =   3960
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblTargetHeartRate 
      BackColor       =   &H00FF8080&
      Caption         =   "Target Heart Rate"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmTargetHeartRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Age As Integer
Dim RHR As Integer
Dim Result1 As Integer
Dim Result2 As Integer
Dim Low As Integer
Dim High As Integer
Dim Min As Integer
Dim Max As Integer


Private Sub Command1_Click()

frmMainPage.Show
frmTargetHeartRate.Hide

End Sub

Private Sub cmdback_Click()
frmMainPage.Show
frmTargetHeartRate.Hide
End Sub

Private Sub cmdCalculate_Click()
Age = txtAge.Text
RHR = txtResting.Text

picResults1.Cls
picResults2.Cls


'220 -age = 197
Result1 = 220 - Age

'197 - resting heart rate = 132
Result2 = Result1 - RHR

'result of previous(132) *65% =(low end of heat rate)
Low = 0.65 * Result2

'result of step 2(132)* 85% =(high end)
High = 0.85 * Result2

'Low end + resting heart rate = min.
Min = Low + RHR

'High end + rhr = max.
Max = High + RHR

'The target heart rate zone for this person would be min to max
picResults1.Print Min
picResults2.Print Max


End Sub

