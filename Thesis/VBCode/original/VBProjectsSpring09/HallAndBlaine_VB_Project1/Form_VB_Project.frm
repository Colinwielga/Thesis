VERSION 5.00
Begin VB.Form frmQuestion1 
   BackColor       =   &H8000000E&
   Caption         =   "Question 1"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFrmQuestion2 
      Caption         =   "On to question 2!"
      Height          =   2175
      Left            =   7680
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox picResultsSum 
      Height          =   975
      Left            =   4800
      ScaleHeight     =   915
      ScaleWidth      =   1755
      TabIndex        =   10
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdPet 
      Caption         =   "Name a common household pet"
      Height          =   1575
      Left            =   4680
      TabIndex        =   9
      Top             =   5760
      Width           =   2175
   End
   Begin VB.PictureBox Picture9 
      Height          =   855
      Left            =   7800
      ScaleHeight     =   795
      ScaleWidth      =   3795
      TabIndex        =   8
      Top             =   9240
      Width           =   3855
   End
   Begin VB.PictureBox Picture8 
      Height          =   1215
      Left            =   7560
      ScaleHeight     =   1155
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   7800
      Width           =   3975
   End
   Begin VB.PictureBox Picture7 
      Height          =   975
      Left            =   7440
      ScaleHeight     =   915
      ScaleWidth      =   3915
      TabIndex        =   6
      Top             =   6120
      Width           =   3975
   End
   Begin VB.PictureBox picResults5 
      Height          =   1095
      Left            =   7440
      ScaleHeight     =   1035
      ScaleWidth      =   4035
      TabIndex        =   5
      Top             =   4320
      Width           =   4095
   End
   Begin VB.PictureBox picResults4 
      Height          =   735
      Left            =   600
      ScaleHeight     =   675
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   9120
      Width           =   3375
   End
   Begin VB.PictureBox picResults3 
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   7680
      Width           =   3255
   End
   Begin VB.PictureBox picResults2 
      Height          =   1095
      Left            =   480
      ScaleHeight     =   1035
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   6120
      Width           =   3735
   End
   Begin VB.PictureBox picResults1 
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1035
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   4200
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   3000
      Picture         =   "Form_VB_Project.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total points"
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   8400
      Width           =   1455
   End
End
Attribute VB_Name = "frmQuestion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFrmQuestion2_Click()
    frmQuestion1.Hide
    frmQuestion2.Show
End Sub

Private Sub cmdPet_Click()

Dim Pet(1 To 10) As String, Value(1 To 10) As Integer, CTR As Integer, Answer As String, X As Integer
Dim Found As Boolean, Strikes As Integer, Total As Integer, Sum As Integer, Remaining As Integer


Open App.Path & "\pets.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Pet(CTR), Value(CTR)
Loop

Do While Strikes < 3 And Total < 5
Answer = InputBox("Enter your answer in all lower case letters please", "Answer!")
Found = False
    Do While ((Not Found) And (X < CTR))
        X = X + 1
        If Answer = Pet(X) Then
            Found = True
                Select Case X
                    Case Is = 1
                        picResults1.Print Pet(X), Value(X)
                    Case Is = 2
                        picResults2.Print Pet(X), Value(X)
                    Case Is = 3
                        picResults3.Print Pet(X), Value(X)
                    Case Is = 4
                        picResults4.Print Pet(X), Value(X)
                    Case Is = 5
                        picResults5.Print Pet(X), Value(X)
                End Select
            Total = Total + 1
            Sum = Value(X) + Sum
            picResultsSum.Cls
            picResultsSum.Print Sum
        End If
    Loop

If (Not Found) Then
    Strikes = Strikes + 1
    Remaining = 3 - Strikes
    MsgBox "Sorry, but that is not one of the answers! You only have " & Remaining & " remaining.", , "Sorry"
    
End If

X = 0

Loop

If Strikes = 3 Then
    MsgBox "You got three strikes :( on to the next round!", , "Failure"
End If

If Total = 5 Then
    MsgBox "Good Work! You got all the answers right! On to the next round!", , "Great Success!"
End If
cmdFrmQuestion2.Visible = True
End Sub
