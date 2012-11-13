VERSION 5.00
Begin VB.Form frmPresidents 
   Caption         =   "CSB/SJU Presidents by year"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   Picture         =   "frmstats.frx":0000
   ScaleHeight     =   6915
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFind 
      Caption         =   "Type in a year to see who was president"
      Height          =   1095
      Left            =   600
      TabIndex        =   4
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton cmdAlphabetical 
      Caption         =   "View List of CSB/SJU Presidents Chronologically"
      Height          =   1095
      Left            =   600
      TabIndex        =   3
      Top             =   6000
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      Height          =   6255
      Left            =   9120
      ScaleHeight     =   6195
      ScaleWidth      =   5715
      TabIndex        =   2
      Top             =   1680
      Width           =   5775
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View List of CSB/SJU Presidents "
      Height          =   1095
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Main Menu"
      Height          =   735
      Left            =   10080
      Picture         =   "frmstats.frx":19F52
      TabIndex        =   0
      Top             =   8280
      Width           =   1815
   End
End
Attribute VB_Name = "frmPresidents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fun with CSB/SJU History!
'frmPresidents
'Audrey Gabe
'Written 3/23/09
'This form sorts or searches an array of presidents, their respective school, and year their term as president started


Option Explicit
Dim Counter As Integer
Dim I As Integer
Dim Results As Double
Dim Presidents(1 To 100) As String
Dim Years(1 To 100) As Integer
Dim School(1 To 100) As String

Private Sub cmdAlphabetical_Click()
'array sorted chronologically

Dim TempName As String, Pass As Integer, Pos As Integer
Open App.Path & "\CSBSJUPresidents.txt" For Input As #1 'Opens array file

Counter = 0
Do While Not EOF(1) 'Reads array
    Counter = Counter + 1
    Input #1, Presidents(Counter), School(Counter), Years(Counter)
Loop
    
I = 0
picResults.Cls
picResults.Print "President"; Tab(40); "School"; Tab(60); "Year"
picResults.Print "************************************************************************************"

For Pass = 1 To Counter - 1 'Sorts array by year in ascending order
    For Pos = 1 To Counter - Pass
        If Years(Pos) > Years(Pos + 1) Then
            TempName = Years(Pos)
            Years(Pos) = Years(Pos + 1)
            Years(Pos + 1) = TempName
            TempName = School(Pos)
            School(Pos) = School(Pos + 1)
            School(Pos + 1) = TempName
            TempName = Presidents(Pos)
            Presidents(Pos) = Presidents(Pos + 1)
            Presidents(Pos + 1) = TempName
        End If
    Next Pos
Next Pass

For I = 1 To Counter
    picResults.Print Presidents(I); Tab(40); School(I); Tab(60); Years(I)
Next I

Close #1

End Sub

Private Sub cmdFind_Click()
Dim Found As Boolean, K As String, I As Integer
'user enters year to see what president started that year

picResults.Cls
Open App.Path & "\CSBSJUPresidents.txt" For Input As #1 'Opens file

Counter = 0
Do While Not EOF(1) 'Reads file
    Counter = Counter + 1
    Input #1, Presidents(Counter), School(Counter), Years(Counter)
Loop

Found = False
I = 0

K = InputBox("Enter a year to find who became president", , "Enter year here") 'User enters a year

Do While (Not Found) And I < Counter 'Searching for a matching year in the array
    I = I + 1
    If (Years(I)) = K Then
        Found = True 'A year matches what the user entered
        picResults.Print "President"; Tab(40); "School"; Tab(60); "Years"
        picResults.Print "***********************************************************************************************"
        picResults.Print Presidents(I); Tab(40); School(I); Tab(60); Years(I)
    End If
Loop
If Not Found Then
    picResults.Print K; " was not found as a start year for a president" 'The year the user entered does not match any in the array
End If

End Sub

Private Sub cmdMenu_Click()
frmPresidents.Hide 'brings user back to menu
frmMenu.Show
End Sub

Private Sub cmdView_Click()
Dim Counter As Integer
Dim Results As Single
Dim Presidents(1 To 100) As String
Dim Years(1 To 100) As Integer
Dim School(1 To 100) As String
'prints array

Counter = 0

Open App.Path & "\CSBSJUPresidents.txt" For Input As #1
picResults.Cls
picResults.Print "President"; Tab(40); "School"; Tab(60); "Year"
picResults.Print "************************************************************************"

Do While Not EOF(1) 'Reads the array into the picture box
    Counter = Counter + 1
    Input #1, Presidents(Counter), School(Counter), Years(Counter)
    picResults.Print Presidents(Counter); Tab(40); School(Counter); Tab(60); Years(Counter)
Loop

Close #1

End Sub



