VERSION 5.00
Begin VB.Form frmMainScreen 
   Caption         =   "Main Screen"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   Picture         =   "frmMainScreen.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnglish 
      Caption         =   "English"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdMath 
      Caption         =   "Math"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmMainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEnglish_Click()

'Takes the user to the English Screen

frmMainScreen.Hide
frmEnglish.Show
End Sub

Private Sub cmdMath_Click()

'Takes the user to the Math Screen

frmMainScreen.Hide
frmMath.Show
End Sub

Private Sub cmdQuit_Click()

'Tells the user good luck with their homework and ends the program

MsgBox ("Good luck with your " & Homework & " hours of homework!")
End
End Sub
