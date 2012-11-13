VERSION 5.00
Begin VB.Form frmpart4 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdgotopics 
      Caption         =   "See Pictures"
      Height          =   975
      Left            =   5160
      TabIndex        =   9
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdsalaries 
      BackColor       =   &H8000000E&
      Caption         =   "Does the average Minnesotan make more than the average Timber Wolf?"
      Height          =   735
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   3495
   End
   Begin VB.CommandButton cmdnumberin 
      BackColor       =   &H8000000E&
      Caption         =   "Number Entered"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtnumberentered 
      Height          =   1095
      Left            =   6720
      TabIndex        =   5
      Top             =   3240
      Width           =   2295
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H8000000E&
      Height          =   4935
      Left            =   240
      ScaleHeight     =   4875
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   1200
      Width           =   5775
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H8000000E&
      Caption         =   "Quit"
      Height          =   975
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton cmdnewgame 
      BackColor       =   &H8000000E&
      Caption         =   "Start A New game"
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   2535
   End
   Begin VB.CommandButton cmdattendance 
      BackColor       =   &H8000000E&
      Caption         =   "Click To Learn About Attendance"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblatten 
      BackStyle       =   0  'Transparent
      Caption         =   "What year do you believe is the highest attendance ever record in wolves history?"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6240
      TabIndex        =   6
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label lbltitle4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fun Facts"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   24
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmpart4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Timberwolves basketball
'frmpart4
'nick thielman
'3/20
'on this form the user is able to see timberwolf attendance throughout the years
'The user is then asked what year they believe had the highest attendance and a message
'appears indicating if they are right or wrong. A list then appears showing in decending
'order the attendance with the year it belongs to. Also a total is kept to find player's
'average salary. The user is asked how much they wish to make and this is compared with
'the average salary and a message appears in the picture box.
Option Explicit
Dim years(1 To 20) As Integer, attendance(1 To 20) As Integer, ctr As Integer
Dim tempattendance As Single, tempyears As Single, j As Integer
Dim pos As Integer, pass As Integer, usersguess As Integer
Private Sub cmdattendance_Click()


ctr = 0
'opens file
Open App.Path & "\attendance.txt" For Input As #1
'puts into an array
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, years(ctr), attendance(ctr)
    
Loop
Close #1

picresults.Cls
'determines the year and attendance in decending order
'prints out attendance with the decending order
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If attendance(pos) < attendance(pos + 1) Then
            tempattendance = attendance(pos)
            attendance(pos) = attendance(pos + 1)
            attendance(pos + 1) = tempattendance
            tempyears = years(pos)
            years(pos) = years(pos + 1)
            years(pos + 1) = tempyears
            End If
    Next pos
Next pass
'prints out headers
picresults.Print "yearly average attendance per game"
picresults.Print "*******************************************"
picresults.Print "  BUT WHAT YEAR IS WHAT "
picresults.Print "  "
For j = 1 To ctr
    picresults.Print Tab(10); attendance(j);
Next j

cmdnumberin.Enabled = True

End Sub


Private Sub cmdgotopics_Click()
frmpart4.Hide
frmpart5.Show
End Sub

Private Sub cmdnumberin_click()
Dim guess As Single

picresults.Cls
'takes the users text as a guess
guess = txtnumberentered.Text

'1989
'lets the user guess what year the most attendance took place
If guess = 1989 Then
    MsgBox "That is Correct!", , ""
ElseIf guess >= 1988 Then
    MsgBox "It actually happened before then"
    picresults.Print "It happened in 1989", , ""
ElseIf guess <= 1990 Then
    MsgBox "It actually happened after then"
    picresults.Print "It happened in 1989", , ""
End If




picresults.Print "*****************************"

'determines the year and attendance in decending order

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If attendance(pos) < attendance(pos + 1) Then
            tempattendance = attendance(pos)
            attendance(pos) = attendance(pos + 1)
            attendance(pos + 1) = tempattendance
            tempyears = years(pos)
            years(pos) = years(pos + 1)
            years(pos + 1) = tempyears
            End If
    Next pos
Next pass

'prints header
picresults.Print "Attendance"; Tab(20); "Year"
'prints the years and attendance
For j = 1 To ctr
    picresults.Print attendance(j); Tab(20); years(j)
Next j


End Sub

Private Sub cmdsalaries_Click()
Dim salary(1 To 20) As Single, runningtotal As Single, average As Single, total As Single
Dim money As Single
'set ctr and runningtotal to zero
ctr = 0
runningtotal = 0

'clear picresults
picresults.Cls
'opens file
Open App.Path & "\salaries.txt" For Input As #1
'puts into an array and adds up total
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, salary(ctr)
    runningtotal = runningtotal + salary(ctr)
Loop
Close #1

'gets average
average = runningtotal / ctr

'gets dollar amount from user
money = InputBox("How much money per year do you wish to make someday?")

'determines if the user wants more or less than the average t-wolf
'and prints out difference in the picture box
If money >= average Then
    picresults.Print "You would make "; FormatCurrency(average - money, 0);
    picresults.Print " more than the average Timberwolf"
ElseIf money < average Then
    picresults.Print "You would make "; FormatCurrency(money - average, 0);
    picresults.Print " less than the average Timberwolf"
End If
    
'average minnesotan income = 54634
'subtracts the minnesotan average from the t-wolves
total = average - 54634
'prints header and how much more a player makes
picresults.Print "*******************************************************"
picresults.Print "The average Timberwolf makes "; FormatCurrency(total, 0)
picresults.Print " more than the average Minnesotan"

End Sub



Private Sub cmdnewgame_Click()
'takes user to next form
'close current form opens next form
frmpart4.Hide
frmpart1.Show
End Sub


Private Sub cmdquit_Click()
'quit
End
End Sub

