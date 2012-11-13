VERSION 5.00
Begin VB.Form frmSit1 
   BackColor       =   &H000000FF&
   Caption         =   "Situation 1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Begin Exercise"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   7815
      Left            =   1800
      Picture         =   "frmSit1.frx":0000
      ScaleHeight     =   7755
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label lblScore 
      BackColor       =   &H000000FF&
      Caption         =   $"frmSit1.frx":1B9CB
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSit1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnter_Click()
    Dim Start, Final, Execu, Deduc, Result As Single
    Do While Start <> 8.8
        Start = InputBox("What is the start value of this vault?", "Start Value")
        If Start <> 8.8 Then
            MsgBox "Your Start Value is wrong, try again", , "Error"
        End If
    Loop
    MsgBox "You got the correct start value.", , "Correct"
    Do While Execu <> 0.3
        Execu = InputBox("How much is the execution deduction for this vault?", "Execution Deductions")
        Select Case Execu
            Case Is < 0.3
                MsgBox "Your deductions are too low, try again.", , "Low"
            Case Is > 0.3
                MsgBox "Your deductions as too high, try again.", , "High"
            Case Else
                MsgBox "You got the correct execution deductions.", , "Correct"
        End Select
    Loop
    Do While Deduc <> 0.6
        Deduc = InputBox("How much is the form deduction for this vault?", "Form Deductions")
        Select Case Deduc
            Case Is < 0.6
                MsgBox "Your deductions are too low, try again.", , "Low"
            Case Is > 0.6
                MsgBox "Your deductions as too high, try again.", , "High"
            Case Else
                MsgBox "You got the correct form deductions.", , "Correct"
        End Select
    Loop
    Do While Final <> 7.9
        Final = InputBox("What is the final score?", "Final Score")
        Select Case Final
            Case Is < 7.9
                MsgBox "The score you have given is too low.  Try again.", , "Low"
            Case Is > 7.9
                MsgBox "The score you have given is too high.  Try again", , "High"
            Case Else
               MsgBox "You have the correct final score!", , "Correct"
        End Select
    Loop
    MsgBox "Great job you have passed the first situation.", , "Finished"
    frmSit1.Hide
    frmIntro.Show
    
End Sub

Private Sub cmdReturn_Click()
    frmSit1.Hide
    frmIntro.Show
End Sub
