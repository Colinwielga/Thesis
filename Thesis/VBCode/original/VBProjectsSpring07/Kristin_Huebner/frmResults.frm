VERSION 5.00
Begin VB.Form frmResults 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Quiz Results"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      Caption         =   "Back to main page"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdReview 
      Caption         =   "Review troublesome works"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      ScaleHeight     =   2055
      ScaleWidth      =   6855
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'this section displays the results of the quiz and gives the user the option to review troublesome works

Private Sub cmdHome_Click() 'back to first options form
    frmResults.Hide
    frmChoose_Test.Show
End Sub

Private Sub cmdQuit_Click() 'ends program
    End
End Sub

Private Sub cmdReview_Click() 'opens review form
    frmResults.Hide
    frmReview.Show
End Sub

Private Sub Form_Activate() 'computes and displays results from quiz
Dim Total As Integer, Percent As Single

picResults.Cls

Total = Correct + Incorrect
Percent = Correct / Total

picResults.Print Usr_Name; " out of "; Total; "questions, you submitted "
picResults.Print Correct; " correct answers"
picResults.Print " and "; Incorrect; " incorrect answers."
picResults.Print "Your overall percent of success was "; FormatPercent(Percent, 1); "."
End Sub

