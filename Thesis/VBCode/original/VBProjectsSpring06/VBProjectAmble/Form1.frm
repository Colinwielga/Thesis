VERSION 5.00
Begin VB.Form frmguess 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Enter"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtguess 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lbljeff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Jeff Amble"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblexplain 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Guess what rank the national champion is.  Enter the rank (1-16) here to see if you're right"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmguess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form enables the user to guess what rank the champion of'
'the tournament is'
Option Explicit
'This button enables the user to go back to the main page'
Private Sub cmdback_Click()
    frmguess.Visible = False
    frmmain.Visible = True
End Sub
'This button enables the user to compute their guess'
Private Sub cmdcompute_Click()
    picresults.Cls
    Dim X As Integer
    X = txtguess
        Select Case X
            Case Is = 1
                picresults.Print "North Carolina (a number one seed) won.  You're Right!"
            Case Else
                picresults.Print "Wrong! Try Again"
        End Select
End Sub


