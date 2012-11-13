VERSION 5.00
Begin VB.Form frmGolf 
   Caption         =   "Golf Challenge"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   Picture         =   "frmGolf.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdMain 
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCalloway 
      Caption         =   "Calloway"
      Height          =   1215
      Left            =   5640
      TabIndex        =   4
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Height          =   1215
      Left            =   1080
      TabIndex        =   3
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmdTitleist 
      Caption         =   "Titleist"
      Height          =   1215
      Left            =   5640
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdNike 
      Caption         =   "Nike"
      Height          =   1215
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label lblCompany 
      BackColor       =   &H80000018&
      Caption         =   " Welcome to The Collegeville Long Drive Competition!            Which brand of golf clubs would you like to use?"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   9735
   End
End
Attribute VB_Name = "frmGolf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Each of these commands will take you to the necessary form where you will be able to hit your long drive
'Except for the cmdMain_click which will take you back to the main menu
Private Sub cmdCalloway_Click()
    frmCalloway.Show
    frmGolf.Hide
End Sub

Private Sub cmdMain_Click()

    frmGolf.Hide
    frmHome.Show
End Sub

Private Sub cmdNike_Click()
    frmNike.Show
    frmGolf.Hide
End Sub

Private Sub cmdPing_Click()
    frmPing.Show
    frmGolf.Hide
End Sub

Private Sub cmdTitleist_Click()
    frmTitleist.Show
    frmGolf.Hide
End Sub
