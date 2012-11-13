VERSION 5.00
Begin VB.Form frmWinners 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBBall 
      Caption         =   "Sort Alphabetically"
      Height          =   615
      Left            =   4320
      TabIndex        =   17
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdhock 
      Caption         =   "Sort Alphabetically"
      Height          =   615
      Left            =   4320
      TabIndex        =   16
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdBaseball 
      Caption         =   "Sort Alphabetically"
      Height          =   615
      Left            =   4320
      TabIndex        =   15
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdfootball 
      Caption         =   "Sort Alphabetically"
      Height          =   615
      Left            =   4320
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Arrays"
      Height          =   975
      Left            =   600
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   975
      Left            =   2040
      TabIndex        =   12
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   7680
      Width           =   1695
   End
   Begin VB.PictureBox picOutput 
      Height          =   10935
      Left            =   5520
      ScaleHeight     =   10875
      ScaleWidth      =   4635
      TabIndex        =   10
      Top             =   0
      Width           =   4695
   End
   Begin VB.TextBox txtbasketball 
      Height          =   615
      Left            =   2040
      TabIndex        =   9
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdBasketball 
      Caption         =   "Find The Winner Of The NBA Finals"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtHockey 
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdHockey 
      Caption         =   "Find The Stanley Cup Winner"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtWorldSeries 
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdWorldSeries 
      Caption         =   "Find The World Series Winner"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtSuperBowl 
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdSuperBowl 
      Caption         =   "Find The Super Bowl Winner"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdB2MM 
      Caption         =   "Go Back to Main Menu"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label lblYear 
      BackColor       =   &H8000000D&
      Caption         =   "Enter The Year You Want To Find The Winner Of"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "frmWinners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdB2MM_Click()
    frmWinners.Hide
    frmMainMenu.Show
End Sub

Private Sub cmdBaseball_Click()
Dim temp As String
Dim pass As Integer
picOutput.Cls
For pass = 1 To Size_WorldSeries - 1
    For A = 1 To Size_WorldSeries - pass
        If WorldSeries(A) > WorldSeries(A + 1) Then
            temp = WorldSeries(A)
            WorldSeries(A) = WorldSeries(A + 1)
            WorldSeries(A + 1) = temp
        End If
    Next A
Next pass

For A = 1 To Size_WorldSeries
    picOutput.Print WorldSeries(A)
Next A
End Sub

Private Sub cmdBasketball_Click()
Dim B, p As Integer
Dim notfound As Boolean
B = txtbasketball.Text
picOutput.Cls
notfound = True
Do While notfound And p < Size_NBA
    p = p + 1
    If B = NBAYear(p) Then
        notfound = False
        picOutput.Print NBA(p); " "; "Won the NBA Finals"; " "; "In"; " "; B
    End If
Loop
    If notfound Then
        picOutput.Print "There was no NBA finals that year"
    End If
End Sub

Private Sub cmdBBall_Click()
Dim temp As String
Dim pass As Integer
picOutput.Cls
For pass = 1 To Size_NBA - 1
    For A = 1 To Size_NBA - pass
        If NBA(A) > NBA(A + 1) Then
            temp = NBA(A)
            NBA(A) = NBA(A + 1)
            NBA(A + 1) = temp
        End If
    Next A
Next pass

For A = 1 To Size_NBA
    picOutput.Print NBA(A)
Next A
End Sub

Private Sub cmdClear_Click()
    picOutput.Cls
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdfootball_Click()
Dim temp As String
Dim pass As Integer
picOutput.Cls
For pass = 1 To Size_SuperBowl - 1
    For A = 1 To Size_SuperBowl - pass
        If SuperBowl(A) > SuperBowl(A + 1) Then
            temp = SuperBowl(A)
            SuperBowl(A) = SuperBowl(A + 1)
            SuperBowl(A + 1) = temp
        End If
    Next A
Next pass

For A = 1 To Size_SuperBowl
    picOutput.Print SuperBowl(A)
Next A
        
End Sub

Private Sub cmdhock_Click()
Dim temp As String
Dim pass As Integer
picOutput.Cls
For pass = 1 To Size_Hockey - 1
    For A = 1 To Size_Hockey - pass
        If Hockey(A) > Hockey(A + 1) Then
            temp = Hockey(A)
            Hockey(A) = Hockey(A + 1)
            Hockey(A + 1) = temp
        End If
    Next A
Next pass

For A = 1 To Size_Hockey
    picOutput.Print Hockey(A)
Next A
End Sub

Private Sub cmdHockey_Click()
Dim C, T As Integer
Dim notfound As Boolean
C = txtHockey.Text
picOutput.Cls
notfound = True
Do While notfound And T < Size_Hockey
T = T + 1
    If C = HockeyYear(T) Then
        notfound = False
        picOutput.Print Hockey(T); " "; "Won the Stanley Cup in"; " "; C
    End If
Loop
    If notfound Then
        picOutput.Print "There was no Stanley Cup winner in that year."
    End If
End Sub

Private Sub cmdLoad_Click()
    Dim I As Integer
    Open App.Path & "\WorldSeries.txt" For Input As #1
     I = 0
    Do Until EOF(1)
            I = I + 1
            Input #1, WSYear(I), WorldSeries(I)
    Loop
    Size_WorldSeries = I
    Close #1
    Open App.Path & "\Superbowl.txt" For Input As #2
    A = 0
    Do Until EOF(2)
        A = A + 1
        Input #2, SBYear(A), SuperBowl(A)
    Loop
    Size_SuperBowl = A
    Close #2
    Open App.Path & "\NBA.txt" For Input As #3
    B = 0
    Do Until EOF(3)
        B = B + 1
        Input #3, NBAYear(B), NBA(B)
    Loop
    Size_NBA = B
    Close #3
    Open App.Path & "\Hockey.txt" For Input As #4
    C = 0
    Do Until EOF(4)
        C = C + 1
        Input #4, HockeyYear(C), Hockey(C)
    Loop
    Size_Hockey = C
    Close #4
End Sub

Private Sub cmdSuperBowl_Click()
Dim K, A As Integer
Dim notfound As Boolean
K = txtSuperBowl.Text
picOutput.Cls
notfound = True
'Project Name: SportsWinners (Rich Muske's SportsWinners.vbp)
'Form Name:  frmWinners (frmWinners.frm)
'Author: Rich Muske
'Date Written: 10/28
'Purpose:  To have someone be able to search for a year and the winner from that major sporting event for that year will show up.  Also to be able to sort the winners of the major sporting events alphabetically for ease to see which team has won a title how many times.



Do While notfound And A < Size_SuperBowl
    A = A + 1
    If K = SBYear(A) Then
        notfound = False
        picOutput.Print SuperBowl(A); " "; "Won the Super Bowl in"; " "; K
    End If
Loop
    If notfound Then
        picOutput.Print "There was no Super Bowl that Year"
    End If
End Sub

Private Sub cmdWorldSeries_Click()
    
    Dim J, I As Integer
    Dim notfound As Boolean
    J = txtWorldSeries.Text
    picOutput.Cls
    notfound = True
    Do While notfound And I < Size_WorldSeries
        I = I + 1
        If J = WSYear(I) Then
            notfound = False
            picOutput.Print WorldSeries(I); " "; "Won the World Series in"; " "; J
        End If
    Loop
    If notfound Then
        picOutput.Print "No World Series That Year"
    End If
End Sub

