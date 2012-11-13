VERSION 5.00
Begin VB.Form Computing_Income 
   BackColor       =   &H00800000&
   Caption         =   "Computing_Income"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdagedif 
      Caption         =   "Do you make more or less than your age group, on average?"
      Height          =   855
      Left            =   480
      TabIndex        =   15
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   2640
      TabIndex        =   14
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton cmddif 
      Caption         =   "Calculate the difference between your earned and expected income."
      Height          =   975
      Left            =   2760
      TabIndex        =   13
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdexpect 
      Caption         =   "Calculate your expected average income."
      Height          =   975
      Left            =   480
      TabIndex        =   12
      Top             =   5880
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFC0C0&
      Height          =   7935
      Left            =   5040
      ScaleHeight     =   7875
      ScaleWidth      =   6075
      TabIndex        =   11
      Top             =   120
      Width           =   6135
   End
   Begin VB.TextBox txtage 
      BackColor       =   &H00FFFFC0&
      Height          =   855
      Left            =   3000
      TabIndex        =   8
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtcollege 
      BackColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   3000
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdincome 
      Caption         =   "Click to enter your income"
      Height          =   735
      Left            =   1560
      TabIndex        =   6
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdmore 
      Caption         =   "6.Average income with more than a bachelors degree"
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdbach 
      Caption         =   "5.Average income with a bachelors degree"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdsome 
      Caption         =   "4.Average Income with some college"
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdnone 
      Caption         =   "3.Average Income with no College"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdsort 
      Caption         =   "2.Sort by highest income"
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdread 
      BackColor       =   &H80000009&
      Caption         =   "1. Read information from file"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Please enter your age in years"
      Height          =   855
      Left            =   600
      TabIndex        =   10
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label lbledu 
      BackColor       =   &H00C0C0FF&
      Caption         =   $"incomevbproject.frx":0000
      Height          =   1095
      Left            =   600
      TabIndex        =   9
      Top             =   3720
      Width           =   2295
   End
End
Attribute VB_Name = "Computing_Income"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Title: Comparing and Computing Income while considering Education and Age
'BY: Erin Granlund
'CS 130, Sprint 2009
' the object is as follows:
' this program determines the averge income of an individual based on there education.
'it also allows the user to put there inforamtion in and do computations for that.


' Declare variable for the form

Dim ctr As Integer
Dim gender(1 To 10) As String, education(1 To 10) As String, income(1 To 10) As String

Dim age(1 To 10) As String
Dim q As Single
Dim e As Single
Dim r As Single
Dim w As Single
Dim j As Single
Dim k As Single
Dim v As Single
Dim b As Single
Dim l As Single
Dim m As Single
Dim averageno As Single
Dim averagesome As Single
Dim averagebach As Single
Dim averagemore As Single
Dim wages As Single
Dim college As Integer
Dim u As Single
Dim t As Single


Dim years As Single























Private Sub cmdbach_Click()
'finds the average income for samples with a bachelors degree

averagebach = w / k
MsgBox "The average income of a sampled individual with a bachelors degree is" & FormatCurrency(averagebach), , "Average Income"

End Sub

Private Sub cmddif_Click()
' finds the difference between earned and expected incomes

picresults.Print
If college = 0 Then
picresults.Print ; "The difference between your expected and earned income is "
picresults.Print ; FormatCurrency(averageno - wages) 'subtracts earned from expected

ElseIf college = 2 Then
picresults.Print ; "The differemce between your expected and earned income is"
picresults.Print ; FormatCurrency(averagesome - wages) 'subtracts earned from expected
ElseIf college = 4 Then
picresults.Print ; "The difference between your expected and earned income is"
picresults.Print ; FormatCurrency(averagebach) 'subtracts earned from expected
ElseIf college = 5 Then
picresults.Print ; "The difference between your expected and earned income is";
picresults.Print ; FormatCurrency(averagemore - wages) 'subtracts earned from expected

End If
' explains the meaning of a negative result

picresults.Print ; "If your difference is negative you earned that much more than expected."


End Sub

Private Sub cmdexpect_Click()
' determines the users expected income

picresults.Print
If college = 0 Then
picresults.Print ; "Your expected average income is "; FormatCurrency(averageno)
ElseIf college = 2 Then
picresults.Print ; "Your expected average income is"; FormatCurrency(averagesome)
ElseIf college = 4 Then
picresults.Print ; "Your expected average income is"; FormatCurrency(averagebach)
ElseIf college = 5 Then
picresults.Print ; "Your expected average income is"; FormatCurrency(averagemore)
End If

End Sub

Private Sub cmdincome_Click()
wages = InputBox("Please enter your earned wages last year rouned to the nearest thousand.")

End Sub

Private Sub cmdmore_Click()
'determines the average income for more than a bachelors degree

averagemore = q / j
MsgBox "The average income of a sampled individual with more than a bachelors degree is" & FormatCurrency(averagemore), , "Average Income"

End Sub

Private Sub cmdname_Click()



End Sub

Private Sub cmdnone_Click()
' finds the average income with no college education
averageno = r / m

MsgBox "The average income of  sampled individual with no college is" & FormatCurrency(averageno), , "Average Income"

End Sub

Private Sub cmdquit_Click()
End

End Sub

Private Sub cmdread_Click()
' set the variable and counter at zero
v = 0
b = 0

 q = 0
w = 0
e = 0
r = 0
j = 0
k = 0
l = 0
m = 0
u = 0
t = 0

 ctr = 0
 ' open the file to read the data
 
Open App.Path & "\vbdata.txt" For Input As #1
Do While Not EOF(1)
ctr = ctr + 1 'set the counter to add one

Input #1, gender(ctr), education(ctr), income(ctr), age(ctr)


' continuously adds the total income for each education set

If education(ctr) = 5 Then
q = q + income(ctr)
ElseIf education(ctr) = 4 Then
w = w + income(ctr)
ElseIf education(ctr) = 2 Then
e = e + income(ctr)
ElseIf education(ctr) = 0 Then
r = r + income(ctr)
End If
' adds the number of data points in each education set

If education(ctr) = 5 Then
j = j + 1
ElseIf education(ctr) = 4 Then
k = k + 1
ElseIf education(ctr) = 2 Then
l = l + 1
ElseIf education(ctr) = 0 Then
m = m + 1
End If

If age(ctr) >= 30 Then
v = v + income(ctr)
t = t + 1
ElseIf age(ctr) <= 29 Then
u = u + 1
b = b + income(ctr)
End If


Loop ' does the process again untill end of file


End Sub

Private Sub cmdsome_Click()
'find the average for samples with some college

averagesome = e / l

MsgBox "The average income for a sampled individual with some college is" & FormatCurrency(averagesome), , "Average Income"

End Sub


Private Sub cmdsort_Click()
'declare the variable

Dim temp As Single
Dim x As String

Dim y As Single

Dim z As Single
Dim pass As Integer, pos As Integer
Dim i As Integer

' sorts the information by highest income
' organized other data points to follow



For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If income(pos) < income(pos + 1) Then
        temp = income(pos)
        income(pos) = income(pos + 1)
        income(pos + 1) = temp
        x = gender(pos)
        gender(pos) = gender(pos + 1)
        gender(pos + 1) = x
        y = age(pos)
        age(pos) = age(pos + 1)
        age(pos + 1) = y
        z = education(pos)
        education(pos) = education(pos + 1)
        education(pos + 1) = z
     End If
    Next pos
Next pass


' prints a header

picresults.Print ; "Income"; Tab(20); "Gender"; Tab(30); "Age"; Tab(40); "Education"

'prints the data
picresults.Print
For i = 1 To ctr

picresults.Print ; FormatCurrency(income(i)); Tab(20); gender(i); Tab(30); age(i); Tab(40); education(i)
Next i

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Command1_Click()
Dim averageover As Single
Dim averageunder As Single

' determine the average for over 30 and under 30 age groups

averageover = v / t
averageunder = b / u


' prints the users results with comprison to these age groups

If years >= 30 Then
picresults.Print
picresults.Print ; "the average for the over 30 age group is"; FormatCurrency(averageover)
picresults.Print ; "The difference between your income and the average for your age is"
picresults.Print ; FormatCurrency(wages - averageover)
ElseIf years <= 29 Then
picresults.Print
picresults.Print ; "The average for the under 30 age group is"; FormatCurrency(averageunder)
picresults.Print ; "The difference between your income and the average for your age is"
picresults.Print ; FormatCurrency(wages - averageunder)
End If

picresults.Print ; "If the differnce is negative then you made that much less than the average"







End Sub

Private Sub txtage_Change()
' the user inputs there age
years = txtage.Text


End Sub

Private Sub txtcollege_Change()
' the user inputs there education

college = txtcollege.Text

End Sub
