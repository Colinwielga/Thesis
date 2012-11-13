VERSION 5.00
Begin VB.Form frmOptionsPage 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Student Options"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   7905
   Begin VB.CommandButton cmdTestFlashCards 
      BackColor       =   &H00000080&
      Caption         =   "Test Flash Cards (NoGrade)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdCreateFlashCards 
      BackColor       =   &H00000080&
      Caption         =   "Create Flash Cards"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdTestVerbs 
      BackColor       =   &H00000080&
      Caption         =   "Test Verb Endings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdViewScores 
      BackColor       =   &H00000080&
      Caption         =   "View Current Scores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdNounTests 
      BackColor       =   &H00000080&
      Caption         =   "Tests Noun Endings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogOut 
      BackColor       =   &H00808080&
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FF0000&
      Caption         =   "Student Options"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   27.75
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   3735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmOptionsPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreateFlashCards_Click()
    'If firstTime Then
        'MsgBox "firsttime = true"
    'Else
        'MsgBox "firstTime = false"
    'End If
    'if the user has never before used the program it creates a new flash card textfile for him/her and gives it an initial entry
    If firstTime Then
        Open App.Path & "\Data\FlashCards\" & userName(StudentPosition) & ".txt" For Append As #1
            Write #1, "This is the First Entry", "It will be ignored", "But it must be present"
        Close #1
    End If
    
    addedFlashVocab = True
    'shows the falshcard creator
    frmOptionsPage.Hide
    frmCreateFlashCards.Show
    'Reads the verbs and nouns just incase
    Call ReadVerbs
    Call ReadNouns
    
End Sub

Private Sub cmdLogOut_Click()
    'Logs out the user
    frmOptionsPage.Hide
    Call LogOut
End Sub

Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub

Private Sub cmdNounTests_Click()
    'Shows the noun test and intiates arrays and useful information
    Dim ctr As Integer
    ctr = 0
    
    'Opens the declension patterns text file and reads it into the arrays(essentialy for tacking on endings)
    Open App.Path & "\Data\DeclensionPatterns.txt" For Input As #1
        Do Until EOF(1)
            ctr = ctr + 1
            Input #1, formName(ctr), First(ctr), SecondM(ctr), SecondN(ctr), ThirdMandF(ctr), ThirdN(ctr), FourthM(ctr), FourthN(ctr), Fifth(ctr)
        Loop
    Close #1
    'Calls the read noun public subroutine, and get class level (cf.mdlpublicsubs)
    Call ReadNouns
    Call GetClassLevel
    'Shows the noun test form
    frmOptionsPage.Hide
    frmTestNouns.Show
End Sub

Private Sub cmdTestFlashCards_Click()
    'Shows the noun test
    frmOptionsPage.Hide
    frmTestFlashCards.Show
End Sub

Private Sub cmdTestVerbs_Click()
    frmVerbTest.Show
    frmOptionsPage.Hide
    
    
    Call ReadVerbs
    Call GetClassLevel
End Sub

Private Sub cmdViewScores_Click()
    frmOptionsPage.Hide
    frmViewScoreData.Show
End Sub
