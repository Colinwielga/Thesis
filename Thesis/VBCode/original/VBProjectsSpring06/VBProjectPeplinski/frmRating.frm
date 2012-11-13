VERSION 5.00
Begin VB.Form frmReview 
   BackColor       =   &H0080FFFF&
   Caption         =   "Review this Program"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   ForeColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturntoMain 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnterData 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save Answers"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtAns9 
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   17
      Top             =   6480
      Width           =   5535
   End
   Begin VB.TextBox txtAns8 
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   15
      Top             =   5640
      Width           =   3975
   End
   Begin VB.TextBox txtAns7 
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   13
      Top             =   5520
      Width           =   3735
   End
   Begin VB.TextBox txtAns6 
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   11
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox txtAns5 
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   9
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox txtAns4 
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox txtAns3 
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtAns2 
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox txtAns1 
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblQues9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "What did the Pattern Test reveal about you?"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Label lblQues8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Write an example of cognitive dissonance"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   14
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label lblQues7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Define cognitive dissonance."
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label lblQues6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "What are the two main ways of impression management?"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   10
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblQues5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Define introspection."
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblQues4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Define self-serving bias."
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblQues3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Why is high self-esteem not always good?"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblQues2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "When are most likely to self-enhance?"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblQues1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Define Implicit Egotism"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEnterData_Click()
'this command button saves answers of the review questions into an array
    
    'declare array and position variable
    Dim AnswerArray(1 To 10) As String
    Dim Pos As Integer
    
    'initializes text box answers into the array
    AnswerArray(1) = txtAns1.Text
    AnswerArray(2) = txtAns2.Text
    AnswerArray(3) = txtAns3.Text
    AnswerArray(4) = txtAns4.Text
    AnswerArray(5) = txtAns5.Text
    AnswerArray(6) = txtAns6.Text
    AnswerArray(7) = txtAns7.Text
    AnswerArray(8) = txtAns8.Text
    AnswerArray(9) = txtAns9.Text
    
    'saves the array in a data file to send to professor as homework
    Open App.Path & "\Homework.txt" For Output As #3
    For Pos = 1 To 9
        Write #3, AnswerArray(Pos)
    Next Pos
    Close #3
    MsgBox "Your answers have been transferred into a data file for your professor to check your work.", , "Message"
    
End Sub
Private Sub cmdReturntoMain_Click()
'Returns user to the main menu
    frmReview.Hide
    frmBegin.Show
End Sub

