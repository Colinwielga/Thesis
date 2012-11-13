VERSION 5.00
Begin VB.Form frmOutOfMoney 
   BackColor       =   &H0000C000&
   Caption         =   "Out of Money"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblOutOfMoney 
      BackColor       =   &H0000C000&
      Caption         =   "I am sorry. Your account has reached a balance of $0.00. Thank you for playing. Please try again later."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2415
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmOutOfMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sports Betting Project
'frmOutOfMoney
'Written by: Sean Egan
'Written on: 3/22/09
'This form is the exit page. When a user's account reaches a balance
' of zero, they are redirected to this form where a label tells
' them that they are out of money and thanks them for playing.

Private Sub cmdExit_Click()
    'Closes the program
    End
End Sub
