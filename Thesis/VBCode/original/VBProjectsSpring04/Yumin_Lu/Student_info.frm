VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   Picture         =   "Student_info.frx":0000
   ScaleHeight     =   6090
   ScaleWidth      =   7170
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000013&
      Height          =   5655
      Left            =   240
      ScaleHeight     =   5595
      ScaleWidth      =   4755
      TabIndex        =   8
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
      Height          =   495
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H80000014&
      Caption         =   "Search by Name"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort by Name"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Set Grade Scale"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdGrade 
      Caption         =   "Show Grade"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Total"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read File"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "YuMin Lu CS130"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   5640
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Grading Program
'Yumin Lu
'March 4th
'Purpose:
'This program will conclude 3 test scores of students,
'including 2 tests and 1 final. It will ask the user for a desired grading scale and print out the grade of each
'student accordingly. The program will also include sorting and searching function.
'The grading method applies the rule that it automatically withdraw the lowest score that is other than the
'final while the final will be double-counted. As well, if the final is the lowest, the total will simply be the
'sum of all three scores.



Option Explicit
Dim Names(1 To 99) As String, Score1(1 To 99) As Single, Score2(1 To 99) As Single, Final(1 To 99) As Single
Dim Total(1 To 99) As Single, Grade(1 To 99) As String 'dim all the varibles
Public CTR As Integer 'dim a count
Public J As Integer
Public Found As Boolean
Public Search As String
Public Position As Integer






Private Sub cmdRead_Click() 'open the file
CTR = 0
Open "info.txt" For Input As #1
picResults.Print "Names", " S1"; "  S2"; "  Final"
picResults.Print "----------------------------------"
Do While Not EOF(1)
CTR = CTR + 1
Input #1, Names(CTR), Score1(CTR), Score2(CTR), Final(CTR)
picResults.Print Names(CTR), Score1(CTR); Score2(CTR); Final(CTR)

Loop

cmdCompute.Enabled = True
cmdRead.Enabled = False



End Sub

Private Sub cmdCompute_Click() 'Compute the total button


picResults.Print "----------------------------------"
picResults.Print "Names", " S1"; "  S2"; "  Final"; " Total"
picResults.Print "----------------------------------"

For J = 1 To CTR

Total(J) = 0
If Final(J) > Score1(J) Then 'If the student's final is not the lowest score, double count it
Total(J) = Total(J) + Final(J) * 2

'If a score that is other than the final is the lowest score, drop it

If Score1(J) > Score2(J) Then
Total(J) = Total(J) + Score1(J)
End If

'Now there are one score and a double-counted final score left
'Add them up, it shall be the total

If Score1(J) < Score2(J) Then
Total(J) = Total(J) + Score2(J)
End If
End If

'if the final is the lowest score
'Simply let the total be the sum of the three scores

If Final(J) < Score1(J) Then
Total(J) = Total(J) + Score1(J)

If Final(J) > Score2(J) Then
Total(J) = Total(J) + Final(J) * 2
End If

If Final(J) < Score2(J) Then
Total(J) = Total(J) + Final(J) + Score2(J)
End If

End If

picResults.Print Names(J), Score1(J); Score2(J); Final(J), Total(J)

'Do the same thing until all students' information has been processed


Next J

'Make the SET GRADING SCALE button available
cmdReset.Enabled = True

End Sub


Private Sub cmdReset_Click()
Form2.Show
Form1.Hide

End Sub

Private Sub cmdGrade_Click()
picResults.Print "----------------------------------"
picResults.Print "Names", "Total", "Grade"
picResults.Print "----------------------------------"


For J = 1 To CTR

If Total(J) >= A Then
'See which category of grade that student's total score fit in
'Assign that grade to the student

picResults.Print Names(J), Total(J), "A"
Grade(J) = "A"
ElseIf Total(J) >= AB Then
picResults.Print Names(J), Total(J), "AB"
Grade(J) = "AB"
ElseIf Total(J) >= B Then
picResults.Print Names(J), Total(J), "B"
Grade(J) = "B"
ElseIf Total(J) >= BC Then
picResults.Print Names(J), Total(J), "BC"
Grade(J) = "BC"
ElseIf Total(J) >= C Then
picResults.Print Names(J), Total(J), "C"
Grade(J) = "C"
ElseIf Total(J) >= CD Then
picResults.Print Names(J), Total(J), "CD"
Grade(J) = "CD"
ElseIf Total(J) >= D Then
picResults.Print Names(J), Total(J), "D"
Grade(J) = "D"
Else
picResults.Print Names(J), Total(J), "F"
Grade(J) = "F"
End If

'Repeat these actions until all students have their grade
Next J

cmdSearch.Enabled = True
cmdSort.Enabled = True

End Sub

Private Sub cmdSearch_Click()

'Get a name from user
Search = InputBox("Enter a name to search for")
Found = False
Position = 0
Do While ((Not Found) And (Position < CTR))
Position = Position + 1
'Compare the name to the first student's name
If Search = Names(Position) Then
Found = True
End If
'If the names aren't the same, compare the next one
'Repeat this comparing action until all students' name have been compared or the name is found
Loop


'If the names are the same, print "Student found", and the student's name, total, and grade
'If none of the names are the same with the requested name, print "Student Not Found"
If Found Then
picResults.Print "----------------------------------"
picResults.Print "Student Found"
picResults.Print Names(Position), Grade(Position)
Else
MsgBox "Sorry,Student Not Found"
End If


End Sub

Private Sub cmdSort_Click()
Dim Pass As Integer
Dim Comp As Integer
Dim Nametemp As String
Dim Score1temp As Single
Dim Score2temp As Single
Dim Finaltemp As Single
Dim Totaltemp As Single
Dim Gradetemp As String


For Pass = 1 To CTR - 1
   For Comp = 1 To CTR - Pass
   
'If the alphabetical order of the two names is wrong, switch them, together with all other student properties, if the order is right, keep it
'Go to the 2nd and 3rd students' names, repeat the comparing alphabetical order action
   
If Names(Comp) > Names(Comp + 1) Then
Nametemp = Names(Comp)
Score1temp = Score1(Comp)
Score2temp = Score2(Comp)
Finaltemp = Final(Comp)
Totaltemp = Total(Comp)
Gradetemp = Grade(Comp)

Names(Comp) = Names(Comp + 1)
Score1(Comp) = Score1(Comp + 1)
Score2(Comp) = Score2(Comp + 1)
Final(Comp) = Final(Comp + 1)
Total(Comp) = Total(Comp + 1)
Grade(Comp) = Grade(Comp + 1)

Names(Comp + 1) = Nametemp
Score1(Comp + 1) = Score1temp
Score2(Comp + 1) = Score2temp
Final(Comp + 1) = Finaltemp
Total(Comp + 1) = Totaltemp
Grade(Comp + 1) = Gradetemp

End If
'Repeat above three steps until the last two students have been processed

Next Comp


Next Pass

picResults.Print "----------------------------------"
picResults.Print "Names", "Total", "Grade"
picResults.Print "----------------------------------"
For J = 1 To CTR
picResults.Print Names(J), Score1(J); Score2(J); Final(J), Total(J)
Next J


End Sub

Private Sub cmdClear_Click()
picResults.Cls
End Sub

Private Sub cmdQuit_Click()
End
End Sub

