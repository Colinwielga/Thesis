VERSION 5.00
Begin VB.Form frmOwnInput 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Input Your Own Numbers"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   Picture         =   "frmOwnInput.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInputStat 
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Click here to input your own numbers"
      Height          =   1095
      Left            =   2640
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Then, click on the button below to input numbers for the calculation"
      Height          =   735
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblDirection 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First, type the statistic you would like to find (i.e. BA, OPS, OBP, SLG)"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmOwnInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bats As Integer
Dim Hits As Integer
Dim Total As Integer
Dim Walks As Integer
Dim HByPitch As Integer
Dim Sac As Integer
Dim Answer As Single
Dim Search As String

'Baseball Batting Statistics
'frmOwnInput
'Aaron Walsh
'March 24, 2009
'This program will figure out various batting statistics like BA, OPS, OBP, and SLG
'by inputting numbers for certain batting catagories by the user

Private Sub cmdBack_Click()
    frmOwnInput.Hide
    frmInitialform.Show
    
End Sub

Private Sub cmdInput_Click()
'this allows the user to enter their own batting numbers and then determines for them
'the desired batting statistic
    Bats = InputBox("Enter number of At Bats", "AB")
    Hits = InputBox("Enter number of Hits", "H")
    Total = InputBox("Enter number of Total Bases", "TB")
    Walks = InputBox("Enter number of Walks", "BB")
    HByPitch = InputBox("Enter number of times Hit by Pitch", "HBP")
    Sac = InputBox("Enter number of Sacrifice Flies", "SF")
    Search = txtInputStat.Text
    If Search = "BA" Then
        Answer = Hits / Bats
        Select Case Answer
            Case 0.3 To 1
                MsgBox "Your Batting Average is " & FormatNumber(Answer, 3) & " ...that is great!", , "Your BA"
            Case 0 To 0.29999999
                MsgBox "Your Batting Average is " & FormatNumber(Answer, 3) & " ...time to hit the batting cages", , "Your BA"
            Case Else
                MsgBox "Error-You need check your inputs", , "Error"
        End Select
    ElseIf Search = "OBP" Then
        Answer = (Hits + Walks + HByPitch) / (Bats + Walks + HByPitch + Sac)
        Select Case Answer
            Case 0.4 To 1
                MsgBox "Your On Base Percentage is " & FormatNumber(Answer, 3) & " ...that is great!", , "Your OBP"
            Case 0 To 0.39999999
                MsgBox "Your On Base Percentage is " & FormatNumber(Answer, 3) & " ...time to hit the batting cages", , "Your OBP"
            Case Else
                MsgBox "Error-You need check your inputs", , "Error"
        End Select
    ElseIf Search = "OPS" Then
        Answer = ((Hits + Walks + HByPitch) / (Bats + Walks + HByPitch + Sac)) + (Total / Bats)
        Select Case Answer
            Case 0.9 To 5
                MsgBox "Your On Base plus Slugging is " & FormatNumber(Answer, 3) & " ...that is great!", , "Your OPS"
            Case 0 To 0.89999999
                MsgBox "Your On Base plus Slugging is " & FormatNumber(Answer, 3) & " ...time to hit the batting cages", , "Your OPS"
            Case Else
                MsgBox "Error-You need check your inputs", , "Error"
        End Select
    ElseIf Search = "SLG" Then
        Answer = (Total / Bats)
        Select Case Answer
            Case 0.65 To 4
                MsgBox "Your Slugging Percentage is " & FormatNumber(Answer, 3) & " ...that is great!", , "Your SLG"
            Case 0 To 0.64999999
                MsgBox "Your Slugging Percentage is " & FormatNumber(Answer, 3) & " ...time to hit the batting cages", , "Your SLG"
            Case Else
                MsgBox "Error-You need check your inputs", , "Error"
        End Select
    Else
        MsgBox "Error-You need to type in the correct letters...ex. BA", , "Error"
    End If
        
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
