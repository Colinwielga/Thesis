VERSION 5.00
Begin VB.Form frmQuit 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Goodbye"
   ClientHeight    =   6855
   ClientLeft      =   2355
   ClientTop       =   2700
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11520
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   5280
      TabIndex        =   1
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox txtThanks 
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Text            =   "Thank you for using this program."
      Top             =   5040
      Width           =   2535
   End
End
Attribute VB_Name = "frmQuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Form_Load()
    Picture = LoadPicture("M:\CS130\miscellaneous\PROJECTS\RugbyPhoto22.jpg")
End Sub

                                                                        'Eric Bendzinski Project 1.vbp
                                                                        'frmQuitForm
                                                                        'Eric Bendzinski
                                                                        'Written 11/1/06 and 11/3/06
