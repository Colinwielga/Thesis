VERSION 5.00
Begin VB.Form frmOne 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Sort Class Information"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Get Student Grades and Class Averages"
      Height          =   1095
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Open"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H008080FF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Find"
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sort"
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title Calculating Final Grades for a High School English Class
'Lynn Paradis
'October 28-30, 2006
'Objectives - The overall objective of this project was to make my
        'room mate, who is student teaching, an easier way of calculating
        'final grades after a project.  I've changed the names of the students
        'to keep the students identities anonymous.
    'This from is to Sort Names by the Alphabet
        'and to Search for Students.
'Subroutines
    'The first subroutine is to find a student in the list by using an Input Box and
        'a Do While Loop. (The array is loaded into this subroutine so that the user
        'can search for a student without running the whole program)
        'There is then an If/Then statement to give the user feedback
        'about what they entered (print the result if found or an error message in
        ' a MsgBox format.
    'The second subroutine loads the array into the program with a Do While Not Loop so
        'the user can sort the list.
    'The third subroutine is a quit button.
    'The fourth subroutine uses the loaded array to sort the list as mentioned above.
        'The program runs a bubble sort function.
Option Explicit


Private Sub cmdFind_Click()
Dim Find As String, I As Integer
Dim Pos As Integer
picResults.Cls

Find = InputBox("Enter Name of Student You Wish to Find.  Enter Last Name, First Name.", "Find Name")
Open App.Path & "\grades.txt" For Input As #1
I = 1
Do While I <= 5 And Find <> Names(I)
    I = I + 1
    Input #1, Names(I), Age(I), Grade1(I), Grade2(I), Grade3(I), Grade4(I), Grade5(I)
Loop
If Find = Names(I) Then
    picResults.Print "Name is "; Find; "  Age is "; Age(I); "Grades are "; Grade1(I); Grade2(I); Grade3(I); Grade4(I); Grade5(I)
Else
    MsgBox "Name Not Found.", , "Error"
End If
Close #1
End Sub

Private Sub cmdOpen_Click()
Dim I As Integer
Open App.Path & "\grades.txt" For Input As #1
Counter = 0
Do While Not EOF(1)
Counter = Counter + 1
Input #1, Names(Counter), Age(Counter), Grade1(Counter), Grade2(Counter), Grade3(Counter), Grade4(Counter), Grade5(Counter)
Loop

Close #1
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSort_Click()
Dim C As Integer, P As Integer, I As Integer, Temp1 As String
Dim N As Integer, Temp2 As Integer, Temp3 As Integer, Temp4 As Integer
Dim Temp5 As Integer, Temp6 As Integer, Temp7 As Integer
picResults.Cls
For P = 1 To Counter - 1
    For I = 1 To Counter - P
       If Names(I) > Names(I + 1) Then
        Temp1 = Names(I)
        Names(I) = Names(I + 1)
        Names(I + 1) = Temp1
        Temp2 = Age(I)
        Age(I) = Age(I + 1)
        Age(I + 1) = Temp2
        Temp3 = Grade1(I)
        Grade1(I) = Grade1(I + 1)
        Grade1(I + 1) = Temp3
        Temp4 = Grade2(I)
        Grade2(I) = Grade2(I + 1)
        Grade2(I + 1) = Temp4
        Temp5 = Grade3(I)
        Grade3(I) = Grade3(I + 1)
        Grade3(I + 1) = Temp5
        Temp6 = Grade4(I)
        Grade4(I) = Grade4(I + 1)
        Grade4(I + 1) = Temp6
        Temp7 = Grade5(I)
        Grade5(I) = Grade5(I + 1)
        Grade5(I + 1) = Temp7
            End If
    Next I
Next P
For C = 1 To Counter
picResults.Print Names(C); Tab(10), Age(C), Grade1(C), Grade2(C), Grade3(C), Grade4(C), Grade5(C)
Next C
cmdSwitch.Visible = True
End Sub

Private Sub cmdSwitch_Click()
frmOne.Hide
frmTwo.Show

End Sub
