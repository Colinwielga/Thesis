VERSION 5.00
Begin VB.Form frmscores 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   5010
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdavgscores 
      Caption         =   "Click here to find avg. score for any team"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdround6 
      Caption         =   "Championship"
      Height          =   735
      Left            =   3120
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdround5 
      Caption         =   "Final Four"
      Height          =   735
      Left            =   4200
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdround4 
      Caption         =   "Elite 8"
      Height          =   735
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdround3 
      Caption         =   "Sweet 16"
      Height          =   735
      Left            =   5280
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdround2 
      Caption         =   "2nd Round"
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdround1 
      Caption         =   "1st Round"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lbljeff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Jeff Amble"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Image imgncaa 
      Height          =   3000
      Left            =   6360
      Picture         =   "frmscores.frx":0000
      Top             =   1800
      Width           =   2445
   End
   Begin VB.Label lblpickdate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on the round of the game you wish to research"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmscores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form lets the user to decide what round they want scores'
'from.  They click on the corresponding button, and type in the'
'team of the score they desire.  Th user can also view an average'
'of any teams scores throughout the tournament'
Option Explicit
'This button enables the user to enter in a team name and view an'
'average of any teams scores throughout the tournament'
Private Sub cmdavgscores_Click()
    Dim txt As String
    Dim N As String
    Dim count As Integer
    Dim score, TempScore As Integer
    Dim pos As Integer
    Dim Games As Integer
    Dim avg As Single
    N = InputBox("Enter the team you are researching", "AVG. Score")
    Open App.Path & "\averagescores.txt" For Input As #1
    Do Until EOF(1)
        Input #1, txt, TempScore
        pos = InStr(LCase(txt), LCase(N))
        If pos <> 0 Then
            score = score + TempScore
            Games = Games + 1
        End If
    Loop
    If Games > 0 Then
        avg = score / Games
        MsgBox FormatNumber(avg), , "AVG. Score"
    Else
        MsgBox "Team not found, Please enter new team name", , "Error"
    End If
    Close #1
End Sub
'This button enables the user to go back to the main form'
Private Sub cmdback_Click()
    frmscores.Visible = False
    frmmain.Visible = True
End Sub
'This button enables the user to enter in a team and find out'
'their round one score'
Private Sub cmdround1_Click()
    Dim txt As String
    Dim pos As Single
    Dim A As String
    A = InputBox("Enter team you are researching", "Score")
    Open App.Path & "\1stroundscores.txt" For Input As #1
    Do Until (EOF(1) Or pos <> 0)
        Input #1, txt
        pos = InStr(LCase(txt), LCase(A))
    Loop
    
    If A = "" Then
    ElseIf pos <> 0 Then
        MsgBox txt, , "Score"
    Else
        MsgBox "Team not found, Please enter new team name", , "Error"
    End If
    Close #1
End Sub
'This button enables the user to enter in a team and find out'
'their round two score'
Private Sub cmdround2_Click()
    Dim txt As String
    Dim pos As Single
    Dim A As String
    A = InputBox("Enter team you are researching", "Score")
    Open App.Path & "\2ndroundscores.txt" For Input As #1
    Do Until (EOF(1) Or pos <> 0)
        Input #1, txt
        pos = InStr(LCase(txt), LCase(A))
    Loop
    
    If A = "" Then
    ElseIf pos <> 0 Then
        MsgBox txt, , "Score"
    Else
        MsgBox "Team not found, Please enter new team name", , "Error"
    End If
    Close #1
End Sub
'This button enables the user to enter in a team and find out'
'their round three score'
Private Sub cmdround3_Click()
    Dim txt As String
    Dim pos As Single
    Dim A As String
    A = InputBox("Enter team you are researching", "Score")
    Open App.Path & "\3rdroundscores.txt" For Input As #1
    Do Until (EOF(1) Or pos <> 0)
        Input #1, txt
        pos = InStr(LCase(txt), LCase(A))
    Loop
    
    If A = "" Then
    ElseIf pos <> 0 Then
        MsgBox txt, , "Score"
    Else
        MsgBox "Team not found, Please enter new team name", , "Error"
    End If
    Close #1
End Sub
'This button enables the user to enter in a team and find out'
'their round four score'
Private Sub cmdround4_Click()
    Dim txt As String
    Dim pos As Single
    Dim A As String
    A = InputBox("Enter team you are researching", "Score")
    Open App.Path & "\4throundscores.txt" For Input As #1
    Do Until (EOF(1) Or pos <> 0)
        Input #1, txt
        pos = InStr(LCase(txt), LCase(A))
    Loop
    
    If A = "" Then
    ElseIf pos <> 0 Then
        MsgBox txt, , "Score"
    Else
        MsgBox "Team not found, Please enter new team name", , "Error"
    End If
    Close #1
End Sub
'This button enables the user to enter in a team and find out'
'their round five score'
Private Sub cmdround5_Click()
    Dim txt As String
    Dim pos As Single
    Dim A As String
    A = InputBox("Enter team you are researching", "Score")
    Open App.Path & "\5throundscores.txt" For Input As #1
    Do Until (EOF(1) Or pos <> 0)
        Input #1, txt
        pos = InStr(LCase(txt), LCase(A))
    Loop
    
    If A = "" Then
    ElseIf pos <> 0 Then
        MsgBox txt, , "Score"
    Else
        MsgBox "Team not found, Please enter new team name", , "Error"
    End If
    Close #1
End Sub
'This button enables the user to enter in a team and find out'
'their round six score'
Private Sub cmdround6_Click()
    Dim txt As String
    Dim pos As Single
    Dim A As String
    A = InputBox("Enter team you are researching", "Score")
    Open App.Path & "\6throundscores.txt" For Input As #1
    Do Until (EOF(1) Or pos <> 0)
        Input #1, txt
        pos = InStr(LCase(txt), LCase(A))
    Loop
    
    If A = "" Then
    ElseIf pos <> 0 Then
        MsgBox txt, , "Score"
    Else
        MsgBox "Team not found, Please enter new team name", , "Error"
    End If
    Close #1
End Sub

