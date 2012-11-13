VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picComm 
      Height          =   1575
      Left            =   240
      Picture         =   "VBNew.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   6000
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "search"
      Height          =   975
      Left            =   3480
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   2295
      Left            =   2280
      ScaleHeight     =   2235
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label lblComm 
      Caption         =   "Communication Club"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdLoad_Click()
Dim D As Integer
Dim M As Single
Dim Balance As Single
picResults.Cls
M = InputBox("Subtract how much money?", "Money") 'ask for amount of money to add or subtract
Balance = 238.34
If Balance > M Then
        Balance = Balance - M
        picResults.Print "The Communication Club has enough money to cover this."
        picResults.Print "The new balance is"; FormatCurrency(M)
    Else
        picResults.Print "The Communication Club does not have enough"
        picResults.Print "money to cover this expense."
End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSearch_Click()
Dim D As Integer
Dim J As Integer
Dim Balance(1 To 31) As Single
D = InputBox("What is the October Date?", "Enter Date") 'ask october date
Do While D > 31
    MsgBox "Sorry, you have entered an invalid date", , "Error"
    D = InputBox("What is the October Date?", "Enter Date") 'ask october date
Loop
Open "M:\CS130\Projects\VB 10-21-02\Commtxt.txt" For Input As #1
For J = 1 To 31
    Input #1, Balance(J)
Next J
picResults.Print "The balance on October"; D; "was "; FormatCurrency(Balance(D)); "."
End Sub


