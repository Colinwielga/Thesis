VERSION 5.00
Begin VB.Form FrmCourses 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14040
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   14040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSeeCourse 
      Caption         =   "See Golf Course"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   5
      Top             =   5880
      Width           =   2175
   End
   Begin VB.PictureBox picCourse 
      Height          =   5535
      Left            =   5160
      ScaleHeight     =   5475
      ScaleWidth      =   7275
      TabIndex        =   4
      Top             =   3840
      Width           =   7335
   End
   Begin VB.CommandButton cmdBackToHome 
      BackColor       =   &H00808000&
      Caption         =   "Back To Home Screen"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      MaskColor       =   &H00808000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9600
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9600
      Width           =   2175
   End
   Begin VB.CommandButton cmdListCourses 
      Caption         =   "List Courses"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.PictureBox picNiceCourses 
      Height          =   3015
      Left            =   5880
      ScaleHeight     =   2955
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "FrmCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'declares variables
Dim CourseName(1 To 10) As String
Dim CourseState(1 To 10) As String
Dim CourseCity(1 To 10) As String
Dim Counter As Integer
Dim GolfCourseName As String
'lists the top rated golf courses and then gives the user the option to type a golf course name into an input box to make the program show a picture of that golf course

'hides the courses form and makes the title screen form visible
Private Sub cmdBackToHome_Click()
    FrmCourses.Hide
    FrmTitle.Show
End Sub


Private Sub cmdListCourses_Click()
    'prints the labels and table headings
    picNiceCourses.Cls
    picNiceCourses.Print "Top 10 Best Golf Courses In U.S. ranked by Golf Link:"
    picNiceCourses.Print ""
    picNiceCourses.Print "Course Name"; Tab(42); "State"; Tab(62); "City"
    picNiceCourses.Print "_______________________________________________________________"
    'opens the file nicegolfcourses from the folder that the VB program is saved in
    Open App.Path & "\NiceGolfCourses.txt" For Input As #3
    Counter = 0
    'goes through the entire file and saves the file into arrays
    'prints out the file into the nicecourses picture box
    Do While Not EOF(3)
        Counter = Counter + 1
        Input #3, CourseName(Counter), CourseState(Counter), CourseCity(Counter)
        picNiceCourses.Print CourseName(Counter); Tab(40); CourseState(Counter); Tab(60); CourseCity(Counter)
    Loop
    Close #3
End Sub

'ends the program
Private Sub cmdQuit_Click()
    End
End Sub

'clears the picture box so that there are not any previously searched pictures when the form is reopened
Public Sub Clearpictures()
    picCourse.Picture = LoadPicture("")
End Sub

Private Sub cmdSeeCourse_Click()
    picCourse.Cls
    'displays input box that asks the user which golf course he or she would like to see a picture of
    GolfCourseName = InputBox("Enter Name of Course You Would Like To Preview.")
    'reads the name from the input box and when the course name matches the name put in the input
    'box, the corresponding picture is shown
    If GolfCourseName = "Bethpage State Golf Course" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\Bethpage State Golf Course.jpg")
    ElseIf GolfCourseName = "Pine Valley Golf Club" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\Pine Valley Golf Club.jpg")
    ElseIf GolfCourseName = "Jefferson Park Golf Course" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\Jefferson Park Golf Course.jpg")
    ElseIf GolfCourseName = "Augusta National Golf Club" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\Augusta National Golf Club.jpg")
    ElseIf GolfCourseName = "Cypress Point Club" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\Cypress Point Club.jpg")
    ElseIf GolfCourseName = "Pebble Beach Golf Links" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\Pebble Beach Golf Links.jpg")
    ElseIf GolfCourseName = "Liberty National" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\Liberty National.jpg")
    ElseIf GolfCourseName = "Leo J. Martin Memorial Golf Course" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\Leo J. Martin Memorial Golf Course.jpg")
    ElseIf GolfCourseName = "Quail Hollow Club" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\Quail Hollow Club.jpg")
    ElseIf GolfCourseName = "The Wynn Golf Club" Then
        picCourse.Cls
        picCourse.Picture = LoadPicture(App.Path & "\The Wynn Golf Club.jpg")
    'if the name that was entered into the input box does not match a picture in the program file a msg box tells the
    'user to look somewhere else becuase the program does not have it
    Else
        picCourse.Cls
        MsgBox ("Sorry, this program doesn't have a picture of " & GolfCourseName & " on file, you'll have to use Google Images.")
    End If
End Sub
