VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H0000C000&
   Caption         =   "Welcome to Teacher Tool!"
   ClientHeight    =   6510
   ClientLeft      =   1290
   ClientTop       =   1080
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   10545
   Begin VB.TextBox txtDesigner 
      BackColor       =   &H00FF00FF&
      Height          =   285
      Left            =   7920
      TabIndex        =   9
      Text            =   "Designed by Amanda Aamodt"
      Top             =   6120
      Width           =   2535
   End
   Begin VB.PictureBox picWelcome1 
      Height          =   1335
      Left            =   600
      Picture         =   "frmFirst.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picWelcome2 
      Height          =   1335
      Left            =   8760
      Picture         =   "frmFirst.frx":0261
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picWelcome3 
      BackColor       =   &H00FF0000&
      Height          =   3255
      Left            =   240
      Picture         =   "frmFirst.frx":04C2
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   3000
      Width           =   3255
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      MaskColor       =   &H00FF00FF&
      TabIndex        =   2
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdTeachersOnly 
      BackColor       =   &H00FF00FF&
      Caption         =   "Mrs. Wright Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      MaskColor       =   &H00FF00FF&
      TabIndex        =   1
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton cmdStudentsAndTeachers 
      BackColor       =   &H00FF00FF&
      Caption         =   "Students and      Mrs. Wright"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      MaskColor       =   &H00FF00FF&
      TabIndex        =   0
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.Label lblTeachers 
      BackColor       =   &H00FF00FF&
      Caption         =   $"frmFirst.frx":5569
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3840
      TabIndex        =   8
      Top             =   3120
      Width           =   6135
   End
   Begin VB.Label lblStudents 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Students: Click on the button that says ""Students and Teachers"" to view your grade, to find useful formulas, and practice quizzes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   7695
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Welcome Mrs. Wright and her Trigonometry Students to Teacher Tool!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        '"TEACHER TOOL FOR MRS. WRIGHT'S TRIGONOMETRY CLASS"
                        'by Amanda Aamodt
                'CSCI 130 Section 03A Fall 2005
                        '1 November 2005
                
'Project Description
    'The world is turning more and more to computer use for anything and everything,
    'especially in the classroom.  I am currently in school to become a high school
    'math teacher, and I predict that I will need to have a more than just a basic
    'understanding of computers to be effective not only in the classroom, but in
    'other aspects of my life.  Something I appreciate in my classes is when teachers
    'have websites or other types of programs that they use as teaching aids.  Sometimes
    'they have websites where you can see what the assignment for the next class period
    'is.  Others have links to useful websites for the class.  And others still, you
    'can view your current grade.  My project intends to incorporate these and other
    'things to act as a "Teacher Tool."
    'I decided to make this program specific to a trigonometry class, as this is the
    'subject I would like to teach.
    'In this project, there are two parts; a section for the teacher and the students,
    'and a section just for the teacher.  The section just for the teacher is protected
    'by a password.  In the section for both teacher and students, one will find the
    'homework assignment for the next class, a practice quiz, the ability to view one's
    'current grade, and a useful page of formulas and other things related to trignometry.
    'The teacher can view the class list and student grades from the part for teachers
    'only.

'Algorithms
    'Algorithm for cmdClassID
        'to sort the list from the highest to the lowest ID number
            'read the array
            'on the first pass, compare the first ID number with the second ID number
            'if the first ID number is greater than the second ID number, swap the numbers
            'also swap the first and last names so they match their ID numbers
            'then compare the second ID number with the third ID number, using the same
            'process as above
            'do this until all the ID numbers have been compared once
            'on the second pass, compare the first ID number with the second using the same
            'process as in the first pass.  continue down the line.
            'continue the process until all 19 passes have been made
            'print the list
            
    
    'Algorithm for cmdSeeGrade
        'to find FinalGrade:
            'add earned points and divide by the total possible points
        'to find the letter grade based on the FinalGrade:
            'multiply the FinalGrade by 100
            'if 94<= Grade <= 100, then the student receives an A
            'if 92<= Grade < 94, then the student receives an A-
            'if 90<= Grade < 92, then the student receives a B+
            'if 84<= Grade < 90, then the student receives a B
            'if 82<= Grade < 84, then the student receives a B-
            'if 80<= Grade < 82, then the student receives a C+
            'if 74<= Grade < 80, then the student receives a C
            'if 72<= Grade < 74, then the student receives a C-
            'if 70<= Grade < 72, then the student receives a D+
            'if 64<= Grade < 70, then the student receives a D
            'if 62<= Grade < 64, then the student receives a D-
            'if 0<= Grade < 62, then the student receives an F
    
    'Algorithm for cmdBest
        'to sort the list from the highest to the lowest ID number
            'read the array
            'on the first pass, compare the first Test Grade with the Test Grade
            'if the first Test Number is less than the second Test Grade, swap the numbers
            'also swap the first and last names so they match their Test Grades
            'then compare the second Test Grade with the third Test Grade, using the same
            'process as above
            'do this until all the Test Grades have been compared once
            'on the second pass, compare the first Test Grade with the second using the same
            'process as in the first pass.  continue down the line.
            'continue the process until all 19 passes have been made
            'print the list
    
    'Diagram of all forms used in the project
             'frmHome       '-> frmStudent  '-> frmFormulas
                                            '-> frmHomework
                                            '-> frmStudentGrades
                                            '-> frmQuiz
                            '-> frmTeacher  '-> frmClassList
                                            '-> frmTeacherGrades
            
'Project Experience
    'This project was trying for me on many levels.  To start, I had difficulty
    'deciding what topic I should use for my project.  Once I decided to make a tool
    'for teachers to use with their students, then I had to decide what I wanted to
    'include.  I thought about things that I would have liked to have access to as
    'a student.  I thought that I would like to be able to see what the assignment
    'for the next class is.  I thought I would like to take practice quizzes.  I
    'thought that it would be nice to be able to access my grade.  Coming up with how
    'to do all of these things with Visual Basic was not all that difficult, though
    'deciding on the best approach was not always the easiest.  I tried different
    'things, such as using text boxes versus using input boxes.
    'I do view this project as a work in progress.  It was more of an experiment to see
    'what i would be able to do.  If I were to actually do something like this with
    'the classes I teach in the future, I would try to come up with more convenient
    'methods, such as using spreadsheets instead of notepads.
    'The best thing about this project for me is that I accomplished everything without
    'the help of a professor or TA.  The only thing that I could not get to work was
    'changing the color of the buttons.
    'The most difficult part of this project for me was the timing.  There were several
    'outside circumstances that prevented me from putting as much effort into this
    'project as I would have liked to.  However, as I said before, I view this project
    'as a sort of experiment for something I could possibly do in the future.
    'This actually happens to be the reason I took this class; I was hoping to learn
    'skills that I would be able to apply in the classroom as a future educator.  It
    'would be very nice to be able to have something like "Teacher Tool" for my
    'future students.  It would make me feel much more organized, and I am sure that
    'students would appreciate it.

Option Explicit
'this part of the program is the home page
'from here, the students and the teacher can decide where they want to go

Private Sub cmdQuit_Click()
    End     'Ends the program
End Sub

Private Sub cmdStudentsAndTeachers_Click()
    'this button is for getting to the page students and the teacher can view
    frmStudent.Visible = True   'makes the "Students and Teacher" form appear
    frmHome.Visible = False     'hides the "Teacher Tool Home" form
    
End Sub

Private Sub cmdTeachersOnly_Click()
    'this button is for getting to the form that can only be seen by a person with the correct password
    Dim X As String     'declaring variables used
    X = InputBox("Please Enter Password", "Password Request")   'produces an Input box to input the password
    If X = "Math Rocks" Then            'if the password "Math Rocks" is correct, the next form becomes visible
        frmTeacher.Visible = True
        frmHome.Visible = False
        Else
            MsgBox "Password Not Valid", , "Incorrect Password"     'if the password is incorrect, a message box pops up telling the user the password is incorrect
    End If
End Sub

