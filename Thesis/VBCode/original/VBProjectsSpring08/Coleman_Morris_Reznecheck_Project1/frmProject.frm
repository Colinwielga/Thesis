VERSION 5.00
Begin VB.Form frmMainpage 
   BackColor       =   &H00008000&
   Caption         =   "Body Mass Index Calculator"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "What is My Goal Protein Intake?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10440
      TabIndex        =   6
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Should I See a Doctor?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7200
      TabIndex        =   5
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton cmdpyramid 
      Caption         =   "Food Pyramid and Ideal Weight"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      TabIndex        =   4
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton cmdBMI 
      Caption         =   "Check Your BMI"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmdTargetHeartRate 
      Caption         =   "Find Your Target Heart Rate"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Shop Our Wholesome Food Store!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Hangman!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   3840
      Picture         =   "frmProject.frx":0000
      Top             =   240
      Width           =   9675
   End
End
Attribute VB_Name = "frmMainpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBMI_Click()

FormBMI.Show
frmMainpage.Hide

End Sub

Private Sub cmdpyramid_Click()
frm1.Show
frmMainpage.Hide
End Sub

Private Sub cmdTargetHeartRate_Click()
frmTargetHeartRate.Show
frmMainpage.Hide
End Sub

Private Sub Command1_Click()
Hangman.Show
frmMainpage.Hide

End Sub

Private Sub Command2_Click()

frmGrocery_Store.Show
frmMainpage.Hide

End Sub

Private Sub Command3_Click()
frmMainpage.Hide
frmsick.Show
End Sub

Private Sub Command4_Click()
frmMainpage.Hide
frmprotein.Show
End Sub

Private Sub Command5_Click()
End
End Sub
