VERSION 5.00
Begin VB.Form frmTargetHR 
   BackColor       =   &H00800000&
   Caption         =   "Your Target Heart Rate"
   ClientHeight    =   8160
   ClientLeft      =   1245
   ClientTop       =   1620
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12660
   Begin VB.PictureBox picResults2 
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
      Height          =   3495
      Left            =   6840
      ScaleHeight     =   3495
      ScaleWidth      =   5655
      TabIndex        =   4
      Top             =   3240
      Width           =   5655
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
      Left            =   9120
      TabIndex        =   3
      Top             =   6840
      Width           =   2175
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
      Left            =   1440
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   6840
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
      Height          =   1095
      Left            =   960
      ScaleHeight     =   1095
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   1920
      Width           =   5415
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Target Heart Rate"
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
      Left            =   5280
      TabIndex        =   0
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Click the button  below to Calculate your THR"
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
      Left            =   3600
      TabIndex        =   8
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   $"frmTargetHR.frx":0000
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
      Height          =   1335
      Left            =   6840
      TabIndex        =   7
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   $"frmTargetHR.frx":0097
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
      Height          =   3255
      Left            =   840
      TabIndex        =   6
      Top             =   3120
      Width           =   5415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Target Heart Rate (THR)"
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
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "frmTargetHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmTargetHR
'Nick Schuster
'March 26, 2008

'This form calculates various target heart rates for the user based on potential exercise intensity levels
Option Explicit
'To return to the previous form
Private Sub cmdBack_Click()
frmTargetHR.Hide
frmCalculate.Show
End Sub

Private Sub cmdCalculate_Click()
Dim MaxHR As Integer
Dim TargetHR As Integer
Dim J As Single

picResults.Cls
picResults2.Cls

MaxHR = 205.8 - (0.685 * Age)   'Calculates the users maximum heart rate via a standard equation
picResults.Print "Your maximum heart rate is"; MaxHR; "bpm"
picResults.Print "with a standard deviation of 6.4 bpm."
picResults.Print "See chart to the right for your THR."
picResults2.Print "Desired Intensity", "Target Heart Rate"
picResults2.Print "*******************************************************************************************"
For J = 0.5 To 0.9 Step 0.05                                        'Uses a For-Next loop to calculate and display a chart of target
    TargetHR = (MaxHR - RestingHeartRate) * J + RestingHeartRate    'heart rates appropriate for the user at different levels of
    picResults2.Print FormatPercent(J, 0), , TargetHR; "bpm"        'exercise intensity by reference to the user's maximum heart rate
Next J
    
End Sub
'To quit the program now
Private Sub cmdQuit_Click()
End
End Sub


