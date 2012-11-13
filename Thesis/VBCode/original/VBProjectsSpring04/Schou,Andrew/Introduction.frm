VERSION 5.00
Begin VB.Form Introduction 
   Caption         =   "Investment Options"
   ClientHeight    =   2910
   ClientLeft      =   3750
   ClientTop       =   3180
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   7320
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdSelf 
      Caption         =   "Enter my own investment information"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdExist 
      Caption         =   "Show me existing corporate bonds"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   $"Introduction.frx":0000
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "Introduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Investments(project1.vbp)
'Introduction(Introduction.frm)
'Author- Andrew Schou
'3/13/04
'The purpose of this form is to introduce the program and to ask the
'user if he or she wants to see current corporate bonds or enter their own investments.

'the buttons on this form will lead the user to either the investment portion of the program
'or the self investment portion on the program.
Private Sub cmdExist_Click()
Introduction.Hide
Existing.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdSelf_Click()
Introduction.Hide
OpeningForm.Show
End Sub
