VERSION 5.00
Begin VB.Form frmPercentage 
   BackColor       =   &H0000FFFF&
   Caption         =   "Shooting Percentage "
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRetry 
      Caption         =   "Try Again."
      Height          =   975
      Left            =   4800
      TabIndex        =   9
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdFindpercent 
      Caption         =   "Find the percentage!"
      Height          =   975
      Left            =   840
      TabIndex        =   8
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to main page!"
      Height          =   975
      Left            =   720
      TabIndex        =   7
      Top             =   4320
      Width           =   2415
   End
   Begin VB.PictureBox picPercentage 
      BackColor       =   &H0000C000&
      FillColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      ScaleHeight     =   1035
      ScaleWidth      =   4275
      TabIndex        =   6
      Top             =   3000
      Width           =   4335
   End
   Begin VB.TextBox txtAttempts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtMakes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblAttempts 
      Caption         =   "Enter the number of shots attempted=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lblMakes 
      Caption         =   "Enter the number of shots made=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblName 
      Caption         =   "Enter the players name=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmPercentage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
frmPercentage.Hide
frmStormBball.Show
End Sub

Private Sub cmdFindpercent_Click()
Dim Name As String
Dim Makes As Integer, Attempts As Integer, Percentage As Single
Name = txtName.Text 'connects name to the name text box
Makes = txtMakes.Text 'connects makes to make textbox
Attempts = txtAttempts.Text 'connects attempts to attempts textbox
Percentage = (Makes / Attempts) 'formula for percentage of shots made

If txtName.Text = "" Or txtMakes.Text = "" Or txtAttempts.Text = "" Then 'If/Then statement to
    MsgBox "I am going to need more info.", , "Hold Up!"                    'stop user if they don't enter all the data
Else: picPercentage.Print Name; " made"; Makes; "out of"; Attempts; "shots"
      picPercentage.Print "His shooting percentage is:"; FormatPercent(Percentage, 2)
    Select Case Percentage 'select case to give feed back on
    Case Is >= 0.6         'what a good shooting percentage is considered
        picPercentage.Print "Incredible Shooting!"
    Case 0.35 To 0.59
        picPercentage.Print "Good Shooting."
    Case Is < 0.35
        picPercentage.Print "You need practice."
End Select
End If
End Sub

Private Sub cmdRetry_Click()
picPercentage.Cls
txtName.Text = ""
txtMakes.Text = ""
txtAttempts.Text = ""
End Sub
