VERSION 5.00
Begin VB.Form frmFinal 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   FillColor       =   &H80000012&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinalGrade 
      BackColor       =   &H0080C0FF&
      Caption         =   "Show me my FINAL Letter Grade according to the Grading Scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit J.K. Grade Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton cmdFinal 
      BackColor       =   &H0080C0FF&
      Caption         =   "Show me my Letter Grades according to the Grading Scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080C0FF&
      Height          =   8535
      Left            =   120
      ScaleHeight     =   8475
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFinal_Click()

picResults.Print "Your Grade", "Letter Grade on Grading Scale"
J = 0
X = 1

For J = 1 To CtrPct
    Do While PctGrade(J) < GradeScale(X)
        X = X + 1
    Loop
    picResults.Print PctGrade(J), GradeScaleLet(X)
    X = 1
Next J


    
    
End Sub

Private Sub cmdFinalGrade_Click()

picResults.Cls

J = 0
X = 1
Grade = 0

For J = 1 To CtrPct
    Grade = Grade + PctGrade(J)
Next J

Grade = (Grade / (CtrPct * 100)) * 100

Do While Grade < GradeScale(X)
    X = X + 1
Loop

picResults.Print "Your Final Average Grade", "Your Final Letter Grade"
picResults.Print Grade, , GradeScaleLet(X)

End Sub

Private Sub cmdQuit_Click()
End
End Sub
