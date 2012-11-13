VERSION 5.00
Begin VB.Form SequentialSearch 
   BackColor       =   &H00008000&
   Caption         =   "Sequential Search"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleMode       =   0  'User
   ScaleWidth      =   4298.425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "switch forms"
      Height          =   855
      Left            =   6120
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindFirstJuneBday 
      Caption         =   "Find first June birthday in the list"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   6120
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   6480
      TabIndex        =   4
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind3LetterNames 
      Caption         =   "Find and list All 3 letter names"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   6120
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Person Closest to your age"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6240
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin VB.CommandButton cmdReadFile 
      Caption         =   "Read file into arrays"
      Height          =   855
      Left            =   6120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "SequentialSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim path As String
'This program gives several samples that demonstrate several things including:
' file input, sequential search using boolean variables, string and date functions
'Note which variables, listed immediately below, are module level or "global" in scope
Dim People(1 To 30) As String
Dim Bday(1 To 30) As Date
Public CTR As Integer

Private Sub cmdReadFile_Click()
'Read the data in the input file and put it into the array

CTR = 0 'the CTR is used to count the entries and for the array subscripts
'first, open the file to be read
'Open "N:\CS130\VB_examples\birthdays.txt" For Input As #1
Open path & "birthdays.txt" For Input As #1
picResults.Print " The Data file contains the following list of names and birthdates:"
picResults.Print "+++++++++++++++++++++++++++++++++++++++++++++++++"

'then, read the data into the array
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, People(CTR), Bday(CTR)
    picResults.Print People(CTR), Bday(CTR)
Loop
Close   'close the file after using it

cmdFind3LetterNames.Enabled = True
cmdReadFile.Enabled = False

picResults.Print
picResults.Print
End Sub

Private Sub cmdFind3LetterNames_Click()
'look for and print the 3-letter names, if any
Dim j As Integer
Dim found As Boolean
found = False

picResults.Print "The three letter names, if any, are: "
picResults.Print "***************************************"

For j = 1 To CTR
    If Len(People(j)) = 3 Then
        picResults.Print People(j)
        found = True
    End If
Next j

If Not found Then
    picResults.Print "Sorry, no 3 letter names are in the list"
End If

'print 3 blank lines
picResults.Print
picResults.Print


cmdFind3LetterNames.Enabled = False
cmdFindFirstJuneBday.Enabled = True

End Sub

Private Sub cmdFindFirstJuneBday_Click()
'This segment of code finds an prints the first June birthday in the list

Dim found As Boolean
Dim placeCtr As Integer         'keeps track of where you are in the list
placeCtr = 0
found = False

picResults.Print " Is there a June birthday?"

'keep looking as long as you have not found what you are looking for and
' you have not reached the end of the array
Do While (Not found) And (placeCtr < CTR)
    placeCtr = placeCtr + 1
    If (Month(Bday(placeCtr)) = 6) Then
        picResults.Print People(placeCtr); ", born on "; Bday(placeCtr);
        found = True
    End If
Loop

If Not found Then
    picResults.Print "Sorry, no June Birthdays!"
Else
    picResults.Print ", is the first June birthday in the list."
End If
cmdFindFirstJuneBday.Enabled = False
cmdFind.Enabled = True

picResults.Print
picResults.Print
End Sub

Private Sub cmdFind_Click()
'This code segment asks the user for a date and displays the closest date in the list.
Dim j As Integer
Dim YourDate As Date
Dim Closest As Date
Dim difference As Single
difference = 1000000000
Closest = #1/1/1900#

YourDate = InputBox("Enter your date of birth")

picResults.Print , "Your Bday is:", YourDate
picResults.Print "Name", "Birthday", "How close to your Bday, in days"
picResults.Print "****************************************************************"
For j = 1 To CTR
    If (Abs(Bday(j) - YourDate)) < difference Then
        Closest = Bday(j)
        difference = (Abs(Bday(j) - YourDate))
    End If
    picResults.Print People(j), Bday(j), (Abs(Bday(j) - YourDate))
Next j
picResults.Print
picResults.Print " The closest birthday to your's is "; Closest; ","
picResults.Print "which is "; difference; " days away from yours."
cmdFind.Enabled = False

End Sub

Private Sub cmdQuit_Click()
End
End Sub


Private Sub Command1_Click()
frmTwo.Show
End Sub

Private Sub Form_Load()
path = "M:\CS130\SEQ\"
End Sub
