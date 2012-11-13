VERSION 5.00
Begin VB.Form Offense 
   BackColor       =   &H00400000&
   Caption         =   "Minnesota Vikings (Offensive Statistics)"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Back to Player Profile Page"
      Height          =   495
      Left            =   240
      Picture         =   "Offense.frx":0000
      TabIndex        =   10
      Top             =   6600
      Width           =   10455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Sort by Receiving TD's"
      Height          =   735
      Left            =   4680
      TabIndex        =   9
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Sort by Yards per Reception"
      Height          =   855
      Left            =   4680
      TabIndex        =   8
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sort by Rushing TD's"
      Height          =   735
      Left            =   2400
      TabIndex        =   7
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Sort by Yards per Carry (AVG)"
      Height          =   855
      Left            =   2400
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Sort by Passing TD's"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sort by QB Rating"
      Height          =   855
      Left            =   120
      Picture         =   "Offense.frx":09E7
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Receiving"
      Height          =   975
      Left            =   4680
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Rushing"
      Height          =   975
      Left            =   2400
      TabIndex        =   2
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Passing"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   2055
   End
   Begin VB.PictureBox results 
      BackColor       =   &H0000FFFF&
      Height          =   3015
      Left            =   2280
      ScaleHeight     =   2955
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
   Begin VB.Image Image2 
      Height          =   2700
      Left            =   120
      Picture         =   "Offense.frx":13CE
      Top             =   240
      Width           =   2025
   End
   Begin VB.Image Image1 
      Height          =   2625
      Left            =   7680
      Picture         =   "Offense.frx":3462
      Top             =   3720
      Width           =   2925
   End
End
Attribute VB_Name = "Offense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Public Path As String

Private Sub Command1_Click()
Dim Player(1 To 2) As String
Dim Completions(1 To 2), Attempts(1 To 2), Yards(1 To 2), Comppercentage(1 To 2), Longest(1 To 2), TDS(1 To 2), Sacks(1 To 2), INTS(1 To 2), Rating(1 To 2) As Integer
Open Path & "Passing Stats.txt" For Input As #1
results.Cls
results.Print Tab(50); "Passing"
results.Print "Viking";
results.Print Tab(25); "CMP";
results.Print Tab(35); "ATT";
results.Print Tab(45); "YDS";
results.Print Tab(55); "COMP%";
results.Print Tab(68); "LNG";
results.Print Tab(78); "TDS";
results.Print Tab(88); "INT";
results.Print Tab(98); "SACK";
results.Print Tab(108); "RATING"
For I = 1 To 2
    Input #1, Player(I), Completions(I), Attempts(I), Yards(I), Comppercentage(I), Longest(I), TDS(I), INTS(I), Sacks(I), Rating(I)
    results.Print Player(I);
    results.Print Tab(25); Completions(I);
    results.Print Tab(35); Attempts(I);
    results.Print Tab(45); Yards(I);
    results.Print Tab(55); Comppercentage(I);
    results.Print Tab(68); Longest(I);
    results.Print Tab(78); TDS(I);
    results.Print Tab(88); Int(I);
    results.Print Tab(98); Sacks(I);
    results.Print Tab(108); Rating(I)
Next I
Close #1
End Sub

Private Sub Command10_Click()
Dim Average(1 To 12), Fumbles(1 To 12), Attempts(1 To 12), Yards(1 To 12), Longest(1 To 12), TDS(1 To 12) As Double
Dim Tempnum As Double
Dim Tempword As String
Dim Pass, N, I As Integer
N = 12
Dim Player(1 To 12) As String
Open Path & "Receiving Stats.txt" For Input As #1

For I = 1 To 12
        Input #1, Player(I), Attempts(I), Yards(I), Average(I), Longest(I), TDS(I), Fumbles(I)
Next I
Close #1
For Pass = 1 To N - 1
    For I = 1 To N - Pass
        If TDS(I) < TDS(I + 1) Then
            Tempword = Player(I + 1)
            Player(I + 1) = Player(I)
            Player(I) = Tempword
            Tempnum = Attempts(I + 1)
            Attempts(I + 1) = Attempts(I)
            Attempts(I) = Tempnum
            Tempnum = Yards(I + 1)
            Yards(I + 1) = Yards(I)
            Yards(I) = Tempnum
            Tempnum = Average(I + 1)
            Average(I + 1) = Average(I)
            Average(I) = Tempnum
            Tempnum = Longest(I + 1)
            Longest(I + 1) = Longest(I)
            Longest(I) = Tempnum
            Tempnum = TDS(I + 1)
            TDS(I + 1) = TDS(I)
            TDS(I) = Tempnum
            Tempnum = Fumbles(I + 1)
            Fumbles(I + 1) = Fumbles(I)
            Fumbles(I) = Tempnum
        End If
    Next I
Next Pass
results.Cls
results.Print Tab(30); "Receiving by Yards per Reception (AVG)"
results.Print "Viking";
results.Print Tab(25); "REC";
results.Print Tab(35); "Yds";
results.Print Tab(45); "AVG";
results.Print Tab(55); "LONG";
results.Print Tab(65); "TDS";
results.Print Tab(75); "FMBL"
For I = 1 To 12
    results.Print Player(I);
    results.Print Tab(25); Attempts(I);
    results.Print Tab(35); Yards(I);
    results.Print Tab(45); Average(I);
    results.Print Tab(55); Longest(I);
    results.Print Tab(65); TDS(I);
    results.Print Tab(75); Fumbles(I)
Next I

End Sub

Private Sub Command2_Click()
Dim Average(1 To 10), Fumbles(1 To 10), Attempts(1 To 10), Yards(1 To 10), Longest(1 To 10), TDS(1 To 10) As Integer
Dim Player(1 To 10) As String
Open Path & "Russing Stats.txt" For Input As #1
results.Cls
results.Print Tab(50); "Rushing"
results.Print "Viking";
results.Print Tab(25); "ATT";
results.Print Tab(35); "Yds";
results.Print Tab(45); "AVG";
results.Print Tab(55); "LONG";
results.Print Tab(65); "TDS";
results.Print Tab(75); "FMBL"
For I = 1 To 10
    Input #1, Player(I), Attempts(I), Yards(I), Average(I), Longest(I), TDS(I), Fumbles(I)
    results.Print Player(I);
    results.Print Tab(25); Attempts(I);
    results.Print Tab(35); Yards(I);
    results.Print Tab(45); Average(I);
    results.Print Tab(55); Longest(I);
    results.Print Tab(65); TDS(I);
    results.Print Tab(75); Fumbles(I)
Next I
Close #1
End Sub

Private Sub Command3_Click()
Dim Average(1 To 12), Fumbles(1 To 12), Attempts(1 To 12), Yards(1 To 12), Longest(1 To 12), TDS(1 To 12) As Integer
Dim Player(1 To 12) As String
Open Path & "Receiving Stats.txt" For Input As #1

results.Cls
results.Print Tab(50); "Receiving"
results.Print "Viking";
results.Print Tab(25); "REC";
results.Print Tab(35); "Yds";
results.Print Tab(45); "AVG";
results.Print Tab(55); "LONG";
results.Print Tab(65); "TDS";
results.Print Tab(75); "FMBL"
For I = 1 To 12
    Input #1, Player(I), Attempts(I), Yards(I), Average(I), Longest(I), TDS(I), Fumbles(I)
    results.Print Player(I);
    results.Print Tab(25); Attempts(I);
    results.Print Tab(35); Yards(I);
    results.Print Tab(45); Average(I);
    results.Print Tab(55); Longest(I);
    results.Print Tab(65); TDS(I);
    results.Print Tab(75); Fumbles(I)
Next I
Close #1
End Sub

Private Sub Command4_Click()
Offense.Hide
Profiles.Show
End Sub

Private Sub Command5_Click()
Dim Player(1 To 2) As String
Dim Completions(1 To 2), Attempts(1 To 2), Yards(1 To 2), Comppercentage(1 To 2), Longest(1 To 2), TDS(1 To 2), Sacks(1 To 2), INTS(1 To 2), Rating(1 To 2) As Integer
Dim Pass, N As Integer
Dim Tempword As String
Dim Tempnum As Integer
Open Path & "Passing Stats.txt" For Input As #1
N = 2
For I = 1 To 2
   Input #1, Player(I), Completions(I), Attempts(I), Yards(I), Comppercentage(I), Longest(I), TDS(I), INTS(I), Sacks(I), Rating(I)
Next I
Close #1

For Pass = 1 To N - 1
    For I = 1 To N - Pass
        If Rating(I) < Rating(I + 1) Then
           Tempword = Player(I + 1)
           Player(I + 1) = Player(I)
           Player(I) = Tempword
           Tempnum = Completions(I + 1)
           Completions(I + 1) = Completions(I)
           Completions(I) = Tempnum
           Tempnum = Attempts(I + 1)
           Attempts(I + 1) = Attempts(I)
           Attempts(I) = Tempnum
           Tempnum = Yards(I + 1)
           Yards(I + 1) = Yards(I)
           Yards(I) = Tempnum
           Tempnum = Comppercentage(I + 1)
           Comppercentage(I + 1) = Comppercentage(I)
           Comppercentage(I) = Tempnum
           Tempnum = Longest(I + 1)
           Longest(I + 1) = Longest(I)
           Longest(I) = Tempnum
           Tempnum = TDS(I + 1)
           TDS(I + 1) = TDS(I)
           TDS(I) = Tempnum
           Tempnum = INTS(I + 1)
           INTS(I + 1) = INTS(I)
           INTS(I) = Tempnum
           Tempnum = Sacks(I + 1)
           Sacks(I + 1) = Sacks(I)
           Sacks(I) = Tempnum
           Tempnum = Rating(I + 1)
           Rating(I + 1) = Rating(I)
           Rating(I) = Tempnum
           Tempnum = INTS(I + 1)
           INTS(I + 1) = INTS(I)
           INTS(I) = Tempnum
        End If
    Next I
Next Pass
results.Cls
results.Print Tab(50); "Passing by QB Rating"
results.Print "Viking";
results.Print Tab(25); "CMP";
results.Print Tab(35); "ATT";
results.Print Tab(45); "YDS";
results.Print Tab(55); "COMP%";
results.Print Tab(68); "LNG";
results.Print Tab(78); "TDS";
results.Print Tab(88); "INT";
results.Print Tab(98); "SACK";
results.Print Tab(108); "RATING"
For I = 1 To 2
    results.Print Player(I);
    results.Print Tab(25); Completions(I);
    results.Print Tab(35); Attempts(I);
    results.Print Tab(45); Yards(I);
    results.Print Tab(55); Comppercentage(I);
    results.Print Tab(68); Longest(I);
    results.Print Tab(78); TDS(I);
    results.Print Tab(88); Int(I);
    results.Print Tab(98); Sacks(I);
    results.Print Tab(108); Rating(I)
Next I


End Sub

Private Sub Command6_Click()
Dim Player(1 To 2) As String
Dim Completions(1 To 2), Attempts(1 To 2), Yards(1 To 2), Comppercentage(1 To 2), Longest(1 To 2), TDS(1 To 2), Sacks(1 To 2), INTS(1 To 2), Rating(1 To 2) As Integer
Dim Pass, N As Integer
Dim Tempword As String
Dim Tempnum As Integer
N = 2
Open Path & "Passing Stats.txt" For Input As #1
For I = 1 To 2
   Input #1, Player(I), Completions(I), Attempts(I), Yards(I), Comppercentage(I), Longest(I), TDS(I), INTS(I), Sacks(I), Rating(I)
Next I
Close #1

For Pass = 1 To N - 1
    For I = 1 To N - Pass
        If TDS(I) < TDS(I + 1) Then
           Tempword = Player(I + 1)
           Player(I + 1) = Player(I)
           Player(I) = Tempword
           Tempnum = Completions(I + 1)
           Completions(I + 1) = Completions(I)
           Completions(I) = Tempnum
           Tempnum = Attempts(I + 1)
           Attempts(I + 1) = Attempts(I)
           Attempts(I) = Tempnum
           Tempnum = Yards(I + 1)
           Yards(I + 1) = Yards(I)
           Yards(I) = Tempnum
           Tempnum = Comppercentage(I + 1)
           Comppercentage(I + 1) = Comppercentage(I)
           Comppercentage(I) = Tempnum
           Tempnum = Longest(I + 1)
           Longest(I + 1) = Longest(I)
           Longest(I) = Tempnum
           Tempnum = TDS(I + 1)
           TDS(I + 1) = TDS(I)
           TDS(I) = Tempnum
           Tempnum = INTS(I + 1)
           INTS(I + 1) = INTS(I)
           INTS(I) = Tempnum
           Tempnum = Sacks(I + 1)
           Sacks(I + 1) = Sacks(I)
           Sacks(I) = Tempnum
           Tempnum = Rating(I + 1)
           Rating(I + 1) = Rating(I)
           Rating(I) = Tempnum
           Tempnum = INTS(I + 1)
           INTS(I + 1) = INTS(I)
           INTS(I) = Tempnum
        End If
    Next I
Next Pass
results.Cls
results.Print Tab(50); "Passing by Touch Downs"
results.Print "Viking";
results.Print Tab(25); "CMP";
results.Print Tab(35); "ATT";
results.Print Tab(45); "YDS";
results.Print Tab(55); "COMP%";
results.Print Tab(68); "LNG";
results.Print Tab(78); "TDS";
results.Print Tab(88); "INT";
results.Print Tab(98); "SACK";
results.Print Tab(108); "RATING"
For I = 1 To 2
    results.Print Player(I);
    results.Print Tab(25); Completions(I);
    results.Print Tab(35); Attempts(I);
    results.Print Tab(45); Yards(I);
    results.Print Tab(55); Comppercentage(I);
    results.Print Tab(68); Longest(I);
    results.Print Tab(78); TDS(I);
    results.Print Tab(88); Int(I);
    results.Print Tab(98); Sacks(I);
    results.Print Tab(108); Rating(I)
Next I
End Sub

Private Sub Command7_Click()
Dim Average(1 To 10), Fumbles(1 To 10), Attempts(1 To 10), Yards(1 To 10), Longest(1 To 10), TDS(1 To 10) As Double
Dim Player(1 To 10) As String
Dim Pass, N As Integer
Dim Tempnum As Double
Dim Tempword As String
N = 10

Open Path & "Russing Stats.txt" For Input As #1

For I = 1 To 10
    Input #1, Player(I), Attempts(I), Yards(I), Average(I), Longest(I), TDS(I), Fumbles(I)
Next I


For Pass = 1 To N - 1
    For I = 1 To N - Pass
        If Average(I) < Average(I + 1) Then
           Tempword = Player(I + 1)
           Player(I + 1) = Player(I)
           Player(I) = Tempword
           Tempnum = Attempts(I + 1)
           Attempts(I + 1) = Attempts(I)
           Attempts(I) = Tempnum
           Tempnum = Yards(I + 1)
           Yards(I + 1) = Yards(I)
           Yards(I) = Tempnum
           Tempnum = Average(I + 1)
           Average(I + 1) = Average(I)
           Average(I) = Tempnum
           Tempnum = Longest(I + 1)
           Longest(I + 1) = Longest(I)
           Longest(I) = Tempnum
           Tempnum = TDS(I + 1)
           TDS(I + 1) = TDS(I)
           TDS(I) = Tempnum
           Tempnum = Fumbles(I + 1)
           Fumbles(I + 1) = Fumbles(I)
           Fumbles(I) = Tempnum
        End If
    Next I
Next Pass

results.Cls
results.Print Tab(30); "Rushing Aranged by Yards per Carry (AVG)"
results.Print "Viking";
results.Print Tab(25); "ATT";
results.Print Tab(35); "Yds";
results.Print Tab(45); "AVG";
results.Print Tab(55); "LONG";
results.Print Tab(65); "TDS";
results.Print Tab(75); "FMBL"
For I = 1 To 10
    results.Print Player(I);
    results.Print Tab(25); Attempts(I);
    results.Print Tab(35); Yards(I);
    results.Print Tab(45); Average(I);
    results.Print Tab(55); Longest(I);
    results.Print Tab(65); TDS(I);
    results.Print Tab(75); Fumbles(I)
Next I
Close #1
End Sub

Private Sub Command8_Click()
Dim Average(1 To 10), Fumbles(1 To 10), Attempts(1 To 10), Yards(1 To 10), Longest(1 To 10), TDS(1 To 10) As Double
Dim Player(1 To 10) As String
Dim Pass, N As Integer
Dim Tempnum As Double
Dim Tempword As String
N = 10

Open Path & "Russing Stats.txt" For Input As #1

For I = 1 To 10
    Input #1, Player(I), Attempts(I), Yards(I), Average(I), Longest(I), TDS(I), Fumbles(I)
Next I


For Pass = 1 To N - 1
    For I = 1 To N - Pass
        If TDS(I) < TDS(I + 1) Then
           Tempword = Player(I + 1)
           Player(I + 1) = Player(I)
           Player(I) = Tempword
           Tempnum = Attempts(I + 1)
           Attempts(I + 1) = Attempts(I)
           Attempts(I) = Tempnum
           Tempnum = Yards(I + 1)
           Yards(I + 1) = Yards(I)
           Yards(I) = Tempnum
           Tempnum = Average(I + 1)
           Average(I + 1) = Average(I)
           Average(I) = Tempnum
           Tempnum = Longest(I + 1)
           Longest(I + 1) = Longest(I)
           Longest(I) = Tempnum
           Tempnum = TDS(I + 1)
           TDS(I + 1) = TDS(I)
           TDS(I) = Tempnum
           Tempnum = Fumbles(I + 1)
           Fumbles(I + 1) = Fumbles(I)
           Fumbles(I) = Tempnum
        End If
    Next I
Next Pass

results.Cls
results.Print Tab(30); "Rushing Aranged by Touch Downs"
results.Print "Viking";
results.Print Tab(25); "ATT";
results.Print Tab(35); "Yds";
results.Print Tab(45); "AVG";
results.Print Tab(55); "LONG";
results.Print Tab(65); "TDS";
results.Print Tab(75); "FMBL"
For I = 1 To 10
    results.Print Player(I);
    results.Print Tab(25); Attempts(I);
    results.Print Tab(35); Yards(I);
    results.Print Tab(45); Average(I);
    results.Print Tab(55); Longest(I);
    results.Print Tab(65); TDS(I);
    results.Print Tab(75); Fumbles(I)
Next I
Close #1
End Sub

Private Sub Command9_Click()

Dim Average(1 To 12), Fumbles(1 To 12), Attempts(1 To 12), Yards(1 To 12), Longest(1 To 12), TDS(1 To 12) As Double
Dim Tempnum As Double
Dim Tempword As String
Dim Pass, N, I As Integer
N = 12
Dim Player(1 To 12) As String
Open Path & "Receiving Stats.txt" For Input As #1

For I = 1 To 12
        Input #1, Player(I), Attempts(I), Yards(I), Average(I), Longest(I), TDS(I), Fumbles(I)
Next I
Close #1
For Pass = 1 To N - 1
    For I = 1 To N - Pass
        If Average(I) < Average(I + 1) Then
            Tempword = Player(I + 1)
            Player(I + 1) = Player(I)
            Player(I) = Tempword
            Tempnum = Attempts(I + 1)
            Attempts(I + 1) = Attempts(I)
            Attempts(I) = Tempnum
            Tempnum = Yards(I + 1)
            Yards(I + 1) = Yards(I)
            Yards(I) = Tempnum
            Tempnum = Average(I + 1)
            Average(I + 1) = Average(I)
            Average(I) = Tempnum
            Tempnum = Longest(I + 1)
            Longest(I + 1) = Longest(I)
            Longest(I) = Tempnum
            Tempnum = TDS(I + 1)
            TDS(I + 1) = TDS(I)
            TDS(I) = Tempnum
            Tempnum = Fumbles(I + 1)
            Fumbles(I + 1) = Fumbles(I)
            Fumbles(I) = Tempnum
        End If
    Next I
Next Pass
results.Cls
results.Print Tab(30); "Receiving by Yards per Reception (AVG)"
results.Print "Viking";
results.Print Tab(25); "REC";
results.Print Tab(35); "Yds";
results.Print Tab(45); "AVG";
results.Print Tab(55); "LONG";
results.Print Tab(65); "TDS";
results.Print Tab(75); "FMBL"
For I = 1 To 12
    results.Print Player(I);
    results.Print Tab(25); Attempts(I);
    results.Print Tab(35); Yards(I);
    results.Print Tab(45); Average(I);
    results.Print Tab(55); Longest(I);
    results.Print Tab(65); TDS(I);
    results.Print Tab(75); Fumbles(I)
Next I

End Sub

Private Sub Form_Load()
Path = "M:\comp. sci\VB Project\"
End Sub
