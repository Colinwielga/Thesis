VERSION 5.00
Begin VB.Form frmCourses 
   BackColor       =   &H0000FF00&
   Caption         =   "Courses"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Go Back to Main Menu"
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCourse3 
      Caption         =   "Blackberry Ridge Golf Course"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   8400
      Width           =   4815
   End
   Begin VB.CommandButton cmdCourse2 
      Caption         =   "Albany Golf Club"
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   4560
      Width           =   4575
   End
   Begin VB.CommandButton cmdCourse1 
      Caption         =   "Rich-Spring Golf Club"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   2745
      Left            =   2760
      Picture         =   "frmCourses.frx":0000
      Top             =   5400
      Width           =   5670
   End
   Begin VB.Image Image1 
      Height          =   2685
      Left            =   5400
      Picture         =   "frmCourses.frx":4391A
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Image imgColdSpring 
      Height          =   2655
      Left            =   240
      Picture         =   "frmCourses.frx":48F84
      Top             =   1560
      Width           =   4875
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: GolfGuide
':Form Name:  frmCourses
':Author:   Tyler Cash
':Date written:  March 19, 2009


'This form displays the three choices for courses.  It allows the user to
'click a button to select a course.
'The program needs to know which course was selected so it knows which pictures to load.

Option Explicit

Private Sub cmdCourse1_Click()
'This button tells the program that the user chose Rich_Spring Golf Club.
    Course = 1
    
'Switching to form displaying specs about selected course
    frmCourseInfo.Show
    frmCourses.Hide
End Sub

Private Sub cmdCourse2_Click()
'This button tells the program that the user chose Albany Golf Club.
    Course = 2
    
 'Switching to form displaying specs about selected course
    frmCourses.Hide
    frmCourseInfo.Show
End Sub

Private Sub cmdCourse3_Click()
'This button tells the program that the user chose Blackberry Ridge Golf Club.
    Course = 3
    
'Switching to form displaying specs about selected course
    frmCourses.Hide
    frmCourseInfo.Show
End Sub

Private Sub cmdExit_Click()
'This button returns the user to the Title Screen.

'Changing forms
    frmTitle.Show
    frmCourses.Hide
End Sub
