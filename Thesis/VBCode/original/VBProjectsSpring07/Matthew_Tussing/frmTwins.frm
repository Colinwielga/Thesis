VERSION 5.00
Begin VB.Form frmTwins 
   BackColor       =   &H000000FF&
   Caption         =   "Twins"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear the Picture Box"
      Height          =   1815
      Left            =   360
      TabIndex        =   6
      Top             =   5400
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      Height          =   6735
      Left            =   6720
      ScaleHeight     =   6675
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1815
      Left            =   3480
      TabIndex        =   4
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdGoback 
      Caption         =   "Go Back To Main Page"
      Height          =   1935
      Left            =   3600
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Click here to see if your favorit player starts for the Minnesota Twins"
      Height          =   1935
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdbatavg 
      Caption         =   "Click here to figure out what you batting average is, or what it would be if you played baseball"
      Height          =   2055
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdRoster 
      Caption         =   "Click here to see the 2007 Twins Starters"
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmTwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numbers(1 To 100) As Integer
Dim twins(1 To 100) As String
Dim ctr As Integer

Private Sub cmdbatavg_Click()
Dim atbats As Integer
Dim hits As Integer
Dim average As Single

    atbats = InputBox("How many at bats did you have", "Input")
    hits = InputBox("How many hits did you get", "input")
    average = hits / atbats
    
    Select Case average
        Case Is > 1
            MsgBox "Sorry you can't have more hits than at bats, try entering your numbers again"
        Case Is >= 0.3
            MsgBox ("Your Batting Average is " & FormatNumber(average, 3) & " You are awsome")
        Case 0.25 To 0.3
            MsgBox ("Your Batting Average is " & FormatNumber(average, 3) & " You are doing good")
        Case 0.2 To 0.25
            MsgBox ("Your Batting Average is " & FormatNumber(average, 3) & " You are doing OK")
        Case Is < 0.2
            MsgBox ("Your Batting Average is " & FormatNumber(average, 3) & " You need to practice")
    End Select

End Sub

Private Sub cmdClear_Click()
    picResults.Cls
End Sub

Private Sub cmdfind_Click()
Dim Found As Boolean
Dim X As Integer, ctr2 As Integer
Dim Name As String

   Open App.Path & "\twins.txt" For Input As #1
    ctr = 0
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, numbers(ctr), twins(ctr)
    Loop
    Close #1
    
    Name = InputBox("Input A Player's Name With The Correct Spelling and Capitalization", "Input")
    
    Found = False
    X = 0
    
    Do While (Found = False And X < ctr)
        X = X + 1
        If twins(X) = Name Then
            Found = True
        End If
    Loop
    
    If Found = True Then
        MsgBox (numbers(X) & " " & twins(X) & " Starts For The Minnesota Twins")
    Else
        MsgBox (Name & " Does not Start For The Minnesota Twins")
    End If
    
End Sub

Private Sub cmdGoback_Click()
    frmProject.Show
    frmTwins.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRoster_Click()
Dim A As Integer
    Open App.Path & "\twins.txt" For Input As #1
    ctr = 0
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, numbers(ctr), twins(ctr)
    Loop
    Close #1
    picResults.Print "The 2007 starters for the Minnesota Twins"
    picResults.Print "---------------------------------------------------------------------"
    For A = 1 To ctr
        picResults.Print numbers(A), twins(A)
    Next A
        
End Sub
