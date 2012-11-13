VERSION 5.00
Begin VB.Form frmStudyGuide 
   BackColor       =   &H0000FFFF&
   Caption         =   "Study Guide"
   ClientHeight    =   11940
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   ScaleHeight     =   11940
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBackWelcome 
      BackColor       =   &H00FF0000&
      Caption         =   "Go Back to Welcome Page"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picWelcome 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   12675
      TabIndex        =   8
      Top             =   2640
      Width           =   12735
   End
   Begin VB.TextBox txtYourAnswer 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1680
      MaxLength       =   60
      TabIndex        =   6
      Top             =   7440
      Width           =   8775
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H00FF0000&
      Caption         =   "Import Study Guide from File"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10920
      Width           =   11895
   End
   Begin VB.CommandButton cmdAnswer 
      BackColor       =   &H00FF0000&
      Caption         =   "Show Answer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdAskQuestion 
      BackColor       =   &H00FF0000&
      Caption         =   "Question"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   12675
      TabIndex        =   0
      Top             =   3840
      Width           =   12735
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FF0000&
      Height          =   3975
      Left            =   9720
      Shape           =   1  'Square
      Top             =   600
      Width           =   3255
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H008080FF&
      Height          =   1455
      Left            =   360
      Shape           =   2  'Oval
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000080&
      Height          =   2895
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   8880
      Width           =   4215
   End
   Begin VB.Shape Shape5 
      Height          =   2895
      Left            =   480
      Shape           =   5  'Rounded Square
      Top             =   2040
      Width           =   6615
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00008080&
      Height          =   8175
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      Height          =   5055
      Left            =   8400
      Shape           =   2  'Oval
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   4575
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   855
      Left            =   4200
      Shape           =   2  'Oval
      Top             =   6360
      Width           =   3855
   End
   Begin VB.Shape shapeTitle 
      BorderColor     =   &H000000FF&
      Height          =   1575
      Left            =   3000
      Shape           =   2  'Oval
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label lblEnterAnswer 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Your Answer Here"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Binary Practice Problems"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   4815
   End
End
Attribute VB_Name = "frmStudyGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declare the global variables for this form
Dim CTR As Single
Dim QuestNum(1 To 100) As Single
Dim Questions(1 To 100) As String

Private Sub cmdGoBackWelcome_Click()
'this button takes the user back to the welcome form
frmWelcome.Show
frmStudyGuide.Hide
frmCompare.Hide
frmAnswer.Hide
End Sub

Private Sub cmdImport_Click()
'these two statements make sure the user selects the correct button first
cmdAnswer.Enabled = True
cmdAskQuestion.Enabled = True
cmdImport.Visible = False
cmdGoBackWelcome.Visible = True
'open this file to get the questions for the user to answer
Open App.Path & "\Questions.txt" For Input As #1
    Do Until EOF(1)
    CTR = CTR + 1
        Input #1, QuestNum(CTR), Questions(CTR)
        'picResults.Print QuestNum(CTR); Questions(CTR)
    Loop
Close #1
'print a statement to inform the user how to start the questions
    picWelcome.Print "When you are ready to begin "; YourName; " please pick a question"
End Sub
Private Sub cmdQuit_Click()
End
End Sub
Private Sub cmdAskQuestion_Click()
Dim Found As Boolean
Dim Pos As Single
picResults.Cls
picWelcome.Cls
PicQuestion = InputBox("Pick Question 1 through 12")
picWelcome.Print "After you enter your answer, "; YourName; ", you can view the correct answer"
picWelcome.Print "to this question by clicking on Show Answer"
Found = False
'statement needed to allow the user to pick between 10 different questions
Do While Found = False And Pos < CTR
    Pos = Pos + 1
        If QuestNum(Pos) = PicQuestion Then
            Found = True
        End If
    Loop
'statement used to print the question the user choose
        If Found = True Then
            picResults.Print QuestNum(Pos), Questions(Pos)
        End If
End Sub
Private Sub cmdAnswer_Click()
'this button is also where the answer variable is declared
YourAnswer = txtYourAnswer.Text
'this button moves you to the answer form
frmStudyGuide.Visible = False
frmAnswer.Visible = True
'picResults.Print Answer
End Sub
