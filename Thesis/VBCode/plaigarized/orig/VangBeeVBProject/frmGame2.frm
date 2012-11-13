VERSION 5.00
Begin VB.Form frmGame2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Guessing Game"
   ClientHeight    =   13320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   Picture         =   "frmGame2.frx":0000
   ScaleHeight     =   13320
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   1095
      Left            =   7680
      TabIndex        =   3
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Index"
      Height          =   1095
      Left            =   7680
      TabIndex        =   2
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton cmdQuess 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quess Genie's Number"
      Height          =   1095
      Left            =   7680
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGame2.frx":7928
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5520
      TabIndex        =   1
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "frmGame2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

frmGameScene.Show
frmGame2.Hide

End Sub

Private Sub cmdQuess_Click()

Dim Mynum As Integer, x As Integer, YourGuess As String

Randomize
Mynum = Int(Rnd * 10) + 1
Do
    x = x + 1
    YourGuess = InputBox("I'm thinking of a number from 1 through 10. What is my number? (3 tries only)")
    If Not IsNumeric(YourGuess) Then
        MsgBox YourGuess & " is not number. Please enter a number."
    ElseIf CInt(YourGuess) > 10 Or CInt(YourGuess) < 1 Then
        MsgBox YourGuess & " is not between 1 and 10. Try Again!"
    ElseIf CInt(YourGuess) = CInt(Mynum) Then
        MsgBox "You got it! The number is " & CInt(Mynum) & "!"
        Exit Sub
    ElseIf CInt(YourGuess) <> CInt(Mynum) Then
        MsgBox "Nope. Try again."
    End If
    
    If x > 2 Then
        MsgBox "The number that I'm thinking is " & CInt(Mynum) & "!"
        Exit Sub
    End If
Loop

End Sub

Private Sub cmdQuit_Click()

End

End Sub
