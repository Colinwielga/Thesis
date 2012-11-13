VERSION 5.00
Begin VB.Form frmMainPage 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   9855
   ClientLeft      =   2565
   ClientTop       =   450
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleMode       =   0  'User
   ScaleWidth      =   15090
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
      Left            =   600
      TabIndex        =   3
      Top             =   6840
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
      Left            =   600
      TabIndex        =   2
      Top             =   4920
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
      Left            =   600
      TabIndex        =   1
      Top             =   3000
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
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   4200
      Picture         =   "frmMainPage.frx":0000
      Top             =   480
      Width           =   9675
   End
End
Attribute VB_Name = "frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBMI_Click()

frmBMI.Show
frmMainPage.Hide


End Sub

Private Sub cmdTargetHeartRate_Click()
frmTargetHeartRate.Show
frmMainPage.Hide
End Sub

Private Sub Command1_Click()
frmHangman1.Show
frmMainPage.Hide

End Sub

Private Sub Command2_Click()

frmGrocery_Store.Show
frmMainPage.Hide

End Sub
