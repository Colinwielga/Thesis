VERSION 5.00
Begin VB.Form frmBMR 
   BackColor       =   &H00800000&
   Caption         =   "Your BMR"
   ClientHeight    =   8280
   ClientLeft      =   2880
   ClientTop       =   1275
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   9045
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
      Height          =   615
      Left            =   2040
      ScaleHeight     =   615
      ScaleWidth      =   5415
      TabIndex        =   3
      Top             =   2400
      Width           =   5415
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
      Left            =   6360
      TabIndex        =   2
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
      Left            =   600
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate BMR"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Click the button  below to Calculate your BMR"
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
      TabIndex        =   7
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "(BMR) "
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
      Height          =   855
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Basal Metabolic Rate "
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
      Height          =   975
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   $"frmBMR.frx":0000
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
      Left            =   2040
      TabIndex        =   4
      Top             =   3240
      Width           =   5415
   End
End
Attribute VB_Name = "frmBMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmBMR
'Nick Schuster
'March 26, 2008

'This form calculates and reports the user's BMR based on a standard equation
Option Explicit
'To return to the previous form
Private Sub cmdBack_Click()
frmBMR.Hide
frmCalculate.Show
End Sub

Private Sub cmdCalculate_Click()
Dim BMR As Single

picResults.Cls

If Gender = "M" Then                                                'The program uses and If-Then statement to determine
    BMR = (9.99 * Weight) + (6.25 * Inches) - (4.92 * Age) + 5      'which equation to use based on the user's gender
ElseIf Gender = "F" Then
    BMR = (9.99 * Weight) + (6.25 * Inches) - (4.92 * Age) - 161
End If

picResults.Print "Your BMR is "; Round(BMR); " Calories per day."   'The program reports the user's BMR

End Sub

Private Sub cmdQuit_Click()
End
End Sub


