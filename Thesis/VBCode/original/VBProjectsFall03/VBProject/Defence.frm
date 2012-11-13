VERSION 5.00
Begin VB.Form Defence 
   BackColor       =   &H00400000&
   Caption         =   "Defensive Statistics"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Back to Player Profile"
      Height          =   1095
      Left            =   8280
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sort by Sacks"
      Height          =   975
      Left            =   8280
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sort by Interseptions"
      Height          =   975
      Left            =   8280
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.PictureBox results 
      BackColor       =   &H0000FFFF&
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4635
      ScaleWidth      =   7875
      TabIndex        =   2
      Top             =   120
      Width           =   7935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Farorite Vikings Defensive Player"
      Height          =   975
      Left            =   8280
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vikings Defensive Statistics"
      Height          =   975
      Left            =   8280
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Defence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Player(1 To 21) As String
Public Path As String
Dim Tackels(1 To 21), Solo(1 To 21), Assisted(1 To 21), Sacks(1 To 21), INTS(1 To 21) As Double

Private Sub Command1_Click()
Open Path & "Defensive.txt" For Input As #1

results.Cls
results.Print "Player";
results.Print Tab(23); "Tackels";
results.Print Tab(34); "Solo";
results.Print Tab(44); "ASST";
results.Print Tab(57); "Sacks";
results.Print Tab(67); "INTS"

For I = 1 To 21
    Input #1, Player(I), Tackels(I), Solo(I), Assisted(I), Sacks(I), INTS(I)
    results.Print Player(I);
    results.Print Tab(23); Tackels(I);
    results.Print Tab(34); Solo(I);
    results.Print Tab(44); Assisted(I);
    results.Print Tab(57); Sacks(I);
    results.Print Tab(67); INTS(I)
Next I
Close #1
End Sub

Private Sub Command2_Click()
Dim iname As String
iname = InputBox("Enter the name of your Favorite Viking. If you only know a number, use the Player Profile screen to find the name (Case Sensitive)", "Search")
Dim found As Boolean
found = False
I = 0
results.Cls

Do Until found Or I = 21
        I = I + 1
    If iname = Player(I) Then
        found = True
    End If

Loop

If found = True Then
    results.Print "Player";
    results.Print Tab(23); "Tackels";
    results.Print Tab(34); "Solo";
    results.Print Tab(44); "ASST";
    results.Print Tab(57); "Sacks";
    results.Print Tab(67); "INTS"
        results.Print Player(I);
        results.Print Tab(23); Tackels(I);
        results.Print Tab(34); Solo(I);
        results.Print Tab(44); Assisted(I);
        results.Print Tab(57); Sacks(I);
        results.Print Tab(67); INTS(I)
ElseIf iname = "Nick Schmitz" Then
    results.Print "WOW! I know he's good, but I didn't think you liked him too!!"
Else
    results.Print "Are you sure "; iname; " is a Minnesota Viking Defensive Player?"
End If

End Sub

Private Sub Command3_Click()
Dim Tempword As String
Dim Tempnum As Double
Dim Pass, N As Integer
N = 21

Open Path & "Defensive.txt" For Input As #1
For I = 1 To 21
    Input #1, Player(I), Tackels(I), Solo(I), Assisted(I), Sacks(I), INTS(I)
Next I
Close #1

For Pass = 1 To N - 1
    For I = 1 To N - Pass
        If INTS(I) < INTS(I + 1) Then
            Tempword = Player(I + 1)
            Player(I + 1) = Player(I)
            Player(I) = Tempword
            Tempnum = Solo(I + 1)
            Solo(I + 1) = Solo(I)
            Solo(I) = Tempnum
            Tempnum = Assisted(I + 1)
            Assisted(I + 1) = Assisted(I)
            Assisted(I) = Tempnum
            Tempnum = Sacks(I + 1)
            Sacks(I + 1) = Sacks(I)
            Sacks(I) = Tempnum
            Tempnum = INTS(I + 1)
            INTS(I + 1) = INTS(I)
            INTS(I) = Tempnum
            Tempnum = Tackels(I + 1)
            Tackels(I + 1) = Tackels(I)
            Tackels(I) = Tempnum
        End If
    Next I
Next Pass

results.Cls
results.Print "Player";
results.Print Tab(23); "Tackels";
results.Print Tab(34); "Solo";
results.Print Tab(44); "ASST";
results.Print Tab(57); "Sacks";
results.Print Tab(67); "INTS"

For I = 1 To 21
    results.Print Player(I);
    results.Print Tab(23); Tackels(I);
    results.Print Tab(34); Solo(I);
    results.Print Tab(44); Assisted(I);
    results.Print Tab(57); Sacks(I);
    results.Print Tab(67); INTS(I)
Next I

End Sub

Private Sub Command4_Click()
Dim Tempword As String
Dim Tempnum As Double
Dim Pass, N As Integer
N = 21

Open Path & "Defensive.txt" For Input As #1
For I = 1 To 21
    Input #1, Player(I), Tackels(I), Solo(I), Assisted(I), Sacks(I), INTS(I)
Next I
Close #1

For Pass = 1 To N - 1
    For I = 1 To N - Pass
        If Sacks(I) < Sacks(I + 1) Then
            Tempword = Player(I + 1)
            Player(I + 1) = Player(I)
            Player(I) = Tempword
            Tempnum = Solo(I + 1)
            Solo(I + 1) = Solo(I)
            Solo(I) = Tempnum
            Tempnum = Assisted(I + 1)
            Assisted(I + 1) = Assisted(I)
            Assisted(I) = Tempnum
            Tempnum = Sacks(I + 1)
            Sacks(I + 1) = Sacks(I)
            Sacks(I) = Tempnum
            Tempnum = INTS(I + 1)
            INTS(I + 1) = INTS(I)
            INTS(I) = Tempnum
            Tempnum = Tackels(I + 1)
            Tackels(I + 1) = Tackels(I)
            Tackels(I) = Tempnum
        End If
    Next I
Next Pass

results.Cls
results.Print "Player";
results.Print Tab(23); "Tackels";
results.Print Tab(34); "Solo";
results.Print Tab(44); "ASST";
results.Print Tab(57); "Sacks";
results.Print Tab(67); "INTS"

For I = 1 To 21
    results.Print Player(I);
    results.Print Tab(23); Tackels(I);
    results.Print Tab(34); Solo(I);
    results.Print Tab(44); Assisted(I);
    results.Print Tab(57); Sacks(I);
    results.Print Tab(67); INTS(I)
Next I
End Sub

Private Sub Command5_Click()
Profiles.Show
Defence.Hide

End Sub
