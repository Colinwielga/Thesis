VERSION 5.00
Begin VB.Form RachelHaney7 
   BackColor       =   &H0000FFFF&
   Caption         =   "RachelHaney7"
   ClientHeight    =   4650
   ClientLeft      =   3255
   ClientTop       =   2880
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   6795
   Begin VB.PictureBox picResults 
      Height          =   735
      Left            =   2400
      ScaleHeight     =   675
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdBudget 
      Caption         =   "Your Budget"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblFinal 
      BackColor       =   &H00FF80FF&
      Caption         =   "So, were you within your budget?  Click on the final button to find out!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "RachelHaney7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RachelHaney7 (RachelHaneyVBProject6.frm)
'Rachel Haney 3/11/04
'This form will tell the user how far from their
'budget they are, whether it is positive or
'negative.  It will also give them a message as to
'whether or not they were within their budget by
'displaying a message box

Private Sub cmdBudget_Click()
    Dim Amount As Single
    Amount = Spend - Total
    If Total < Spend Then
            MsgBox "Congratulations!  You are within your budget!  Your trip is a success!", , "Congratulations!"
            picResults.Print "You are within your budget by "; FormatCurrency(Amount); "."
        ElseIf Total > Spend Then
            MsgBox "Sorry.  You are over your planned budget.  Please start over and try again.", , "Sorry"
            picResults.Print "You are over your budget by "; FormatCurrency(Amount); "."
    End If
    cmdBudget.Visible = False
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
