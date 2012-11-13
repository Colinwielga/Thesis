VERSION 5.00
Begin VB.Form frmStudent 
   BackColor       =   &H00C0C000&
   Caption         =   "Students and Teachers"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form2"
   ScaleHeight     =   6255
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesigner 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   5880
      TabIndex        =   6
      Text            =   "Designed by Amanda Aamodt"
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdHomework 
      Caption         =   "What's the Homework???"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdFormulas 
      Caption         =   "Formulas and Some Other Stuff You Should Know"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdStudentGrade 
      Caption         =   "Check Out Your Current Grade"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdPracticeTest 
      Caption         =   "See How Well You're Doing: Take a Practice Test"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.PictureBox picStudentsAndTeachers 
      BackColor       =   &H00FFC0C0&
      Height          =   4815
      Left            =   3240
      Picture         =   "frmSecond.frx":0000
      ScaleHeight     =   4755
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.CommandButton cmdHome1 
      Caption         =   "Back to Home Page"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   0
      Top             =   5280
      Width           =   2175
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this part of the program is the student page
'from here, the student (and the teacher) can decide where they want to go

Private Sub cmdFormulas_Click()
    frmFormulas.Visible = True      'makes the form with the formulas visible
    
End Sub

Private Sub cmdHome1_Click()
    frmStudent.Visible = False      'brings the user back to the home page
    frmHome.Visible = True
End Sub

Private Sub cmdHomework_Click()
    frmHomework.Visible = True      'makes the form with the homework assignment visible
End Sub

Private Sub cmdPracticeTest_Click()
    frmQuiz.Visible = True          'makes the form with the practice quiz visible
End Sub

Private Sub cmdStudentGrade_Click()
    frmStudentGrades.Visible = True 'makes the form for students to view their grades visible
   
End Sub
