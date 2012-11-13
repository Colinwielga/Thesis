VERSION 5.00
Begin VB.Form Profiles 
   BackColor       =   &H00400000&
   Caption         =   "Minnesota Vikings (Player Profiles)"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form 1"
   ScaleHeight     =   8355
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Favorite Player Profile (Enter a Number)"
      Height          =   735
      Left            =   7200
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Defensive Statistics"
      Height          =   735
      Left            =   8520
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort First Half by Number"
      Height          =   735
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Hit the Showers"
      Height          =   1815
      Left            =   9360
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Second Half of Viking's Team (Alphabetically by last name)"
      Height          =   975
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sort Second Half by Number"
      Height          =   735
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox results 
      BackColor       =   &H0000FFFF&
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5835
      ScaleWidth      =   8715
      TabIndex        =   3
      Top             =   2400
      Width           =   8775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Offensive Statistics"
      Height          =   735
      Left            =   9720
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Favorite Player Profile (Enter a Name)"
      Height          =   735
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "First Half of Viking's Team (Alphabetically by last name)"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Head Coach Mike Tice"
      Height          =   735
      Left            =   3600
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   1335
      Left            =   5400
      Picture         =   "Form1.frx":0000
      Top             =   960
      Width           =   1185
   End
   Begin VB.Image Image3 
      Height          =   1080
      Left            =   7320
      Picture         =   "Form1.frx":0904
      Top             =   1200
      Width           =   2475
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   720
      Picture         =   "Form1.frx":1319
      Top             =   1200
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   9000
      Picture         =   "Form1.frx":1D00
      Top             =   2520
      Width           =   2025
   End
End
Attribute VB_Name = "Profiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Tempword, Player(1 To 56), position(1 To 56), length(1 To 56), school(1 To 56), city(1 To 56) As String
Dim Jersey(1 To 56), weight(1 To 56), Pass, Tempnum As Integer
Public Path As String

Private Sub Command1_Click()

Open Path & "Player Profiles.txt" For Input As #1

results.Cls
results.Print Tab(1); "Num.";
results.Print Tab(8); "Player";
results.Print Tab(31); "Pos.";
results.Print Tab(42); "Hgt.";
results.Print Tab(52); "Wgt.";
results.Print Tab(65); "School";
results.Print Tab(88); "City"

For I = 1 To 56
    Input #1, Jersey(I), Player(I), position(I), length(I), weight(I), school(I), city(I)
Next I
For I = 1 To 28
results.Print Tab(1); Jersey(I);
results.Print Tab(8); Player(I);
results.Print Tab(31); position(I);
results.Print Tab(42); length(I);
results.Print Tab(52); weight(I);
results.Print Tab(65); school(I);
results.Print Tab(88); city(I)

Next I
Close #1
End Sub

Private Sub Command2_Click()
Dim found As Boolean
Dim p As String
p = InputBox("Enter your Favorite Minnesota Viking (Case Sensitive)", "Search")
found = False
results.Cls
I = 0

Do Until found Or I = 56
    I = I + 1
    If p = Player(I) Then
        found = True
    End If
Loop
If found Then
    results.Print Tab(1); "Player";
    results.Print Tab(20); "Num.";
    results.Print Tab(31); "Pos.";
    results.Print Tab(42); "Hgt.";
    results.Print Tab(52); "Wgt.";
    results.Print Tab(65); "School";
    results.Print Tab(88); "City"
        results.Print Tab(1); Player(I);
        results.Print Tab(20); Jersey(I);
        results.Print Tab(31); position(I);
        results.Print Tab(42); length(I);
        results.Print Tab(52); weight(I);
        results.Print Tab(65); school(I);
        results.Print Tab(88); city(I)
  ElseIf p = "Nick Schmitz" Then
    results.Print "If Nick Schmitz played for the Vikings, and he could,  he would be my favorite player as well!!"
  Else
    results.Print "Are you sure "; p; " is a Minnesota Viking?"
End If
End Sub

Private Sub Command3_Click()
Offense.Show
Profiles.Hide
End Sub

Private Sub Command4_Click()
Dim N As Integer
N = 56
Open Path & "Player Profiles.txt" For Input As #1

For I = 1 To 56
    Input #1, Jersey(I), Player(I), position(I), length(I), weight(I), school(I), city(I)
Next I
results.Cls
For Pass = 1 To N - 1
    For I = 26 To N - Pass
        If Jersey(I) > Jersey(I + 1) Then
            Tempnum = Jersey(I + 1)
            Jersey(I + 1) = Jersey(I)
            Jersey(I) = Tempnum
            Tempword = Player(I + 1)
            Player(I + 1) = Player(I)
            Player(I) = Tempword
            Tempword = position(I + 1)
            position(I + 1) = position(I)
            position(I) = Tempword
            Tempword = length(I + 1)
            length(I + 1) = length(I)
            length(I) = Tempword
            Tempnum = weight(I + 1)
            weight(I + 1) = weight(I)
            weight(I) = Tempnum
            Tempword = school(I + 1)
            school(I + 1) = school(I)
            school(I) = Tempword
            Tempword = city(I + 1)
            city(I + 1) = city(I)
            city(I) = Tempword
        End If
    Next I
Next Pass

results.Print Tab(1); "Num.";
results.Print Tab(8); "Player";
results.Print Tab(31); "Pos.";
results.Print Tab(42); "Hgt.";
results.Print Tab(52); "Wgt.";
results.Print Tab(65); "School";
results.Print Tab(88); "City"
For I = 29 To 56
results.Print Tab(1); Jersey(I);
results.Print Tab(8); Player(I);
results.Print Tab(31); position(I);
results.Print Tab(42); length(I);
results.Print Tab(52); weight(I);
results.Print Tab(65); school(I);
results.Print Tab(88); city(I)
Next I
Close #1
End Sub

Private Sub Command5_Click()
Open Path & "Player Profiles.txt" For Input As #1
For I = 1 To 56
    Input #1, Jersey(I), Player(I), position(I), length(I), weight(I), school(I), city(I)
Next I

results.Cls

results.Print Tab(1); "Num.";
results.Print Tab(8); "Player";
results.Print Tab(31); "Pos.";
results.Print Tab(42); "Hgt.";
results.Print Tab(52); "Wgt.";
results.Print Tab(65); "School";
results.Print Tab(88); "City"

For I = 29 To 56
results.Print Tab(1); Jersey(I);
results.Print Tab(8); Player(I);
results.Print Tab(31); position(I);
results.Print Tab(42); length(I);
results.Print Tab(52); weight(I);
results.Print Tab(65); school(I);
results.Print Tab(88); city(I)

Next I
Close #1
End Sub

Private Sub Command6_Click()
Dim N As Integer
N = 28
Open Path & "Player Profiles.txt" For Input As #1
For I = 1 To 28
    Input #1, Jersey(I), Player(I), position(I), length(I), weight(I), school(I), city(I)
Next I
results.Cls
For Pass = 1 To N - 1
    For I = 1 To N - Pass
        If Jersey(I) > Jersey(I + 1) Then
            Tempnum = Jersey(I + 1)
            Jersey(I + 1) = Jersey(I)
            Jersey(I) = Tempnum
            Tempword = Player(I + 1)
            Player(I + 1) = Player(I)
            Player(I) = Tempword
            Tempword = position(I + 1)
            position(I + 1) = position(I)
            position(I) = Tempword
            Tempword = length(I + 1)
            length(I + 1) = length(I)
            length(I) = Tempword
            Tempnum = weight(I + 1)
            weight(I + 1) = weight(I)
            weight(I) = Tempnum
            Tempword = school(I + 1)
            school(I + 1) = school(I)
            school(I) = Tempword
            Tempword = city(I + 1)
            city(I + 1) = city(I)
            city(I) = Tempword
        End If
    Next I
Next Pass

results.Print Tab(1); "Num.";
results.Print Tab(8); "Player";
results.Print Tab(31); "Pos.";
results.Print Tab(42); "Hgt.";
results.Print Tab(52); "Wgt.";
results.Print Tab(65); "School";
results.Print Tab(88); "City"
For I = 1 To 28
results.Print Tab(1); Jersey(I);
results.Print Tab(8); Player(I);
results.Print Tab(31); position(I);
results.Print Tab(42); length(I);
results.Print Tab(52); weight(I);
results.Print Tab(65); school(I);
results.Print Tab(88); city(I)
Next I
Close #1
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Command8_Click()
Defence.Show
Profiles.Hide
End Sub

Private Sub Command9_Click()
Dim found As Boolean
Dim iNum As Integer
iNum = InputBox("Enter the Number of your Favorite Minnesota Viking", "Search")
found = False

results.Cls

Open Path & "Player Profiles.txt" For Input As #1
For I = 1 To 56
    Input #1, Jersey(I), Player(I), position(I), length(I), weight(I), school(I), city(I)
Next I
I = 0
Do Until found Or I = 56
    I = I + 1
    If iNum = Jersey(I) Then
        found = True
    End If

Loop

If found Then
    results.Print Tab(1); "Num.";
    results.Print Tab(8); "Player";
    results.Print Tab(31); "Pos.";
    results.Print Tab(42); "Hgt.";
    results.Print Tab(52); "Wgt.";
    results.Print Tab(65); "School";
    results.Print Tab(88); "City"
        results.Print Tab(1); Jersey(I);
        results.Print Tab(8); Player(I);
        results.Print Tab(31); position(I);
        results.Print Tab(42); length(I);
        results.Print Tab(52); weight(I);
        results.Print Tab(65); school(I);
        results.Print Tab(88); city(I)
Else
    results.Print "Are you sure "; iNum; " is a Minnesota Viking?"
End If
Close #1
End Sub

