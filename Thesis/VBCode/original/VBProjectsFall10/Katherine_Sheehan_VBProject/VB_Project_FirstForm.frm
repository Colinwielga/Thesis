VERSION 5.00
Begin VB.Form frmIntroduction 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   5880
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdSwitchForms 
      Caption         =   "Click here to continue to the next page"
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label lblName 
      Caption         =   "By Katherine Sheehan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      Caption         =   "Organizing Sources"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblIntroduction 
      Caption         =   $"VB_Project_FirstForm.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   7455
   End
End
Attribute VB_Name = "frmIntroduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form introduces the user to the project

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSwitchForms_Click()
    frmIntroduction.Hide
    frmReadFiles.Show
    frmCommon.Hide
    frmChicago.Hide
    frmBibliography.Hide
    'the introduction form will be hidden and the next form will be shown.
End Sub
