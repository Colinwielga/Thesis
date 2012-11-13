VERSION 5.00
Begin VB.Form FrmForm4 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form4"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14310
   LinkTopic       =   "Form4"
   Picture         =   "FrmBeauty.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdmenu 
      Caption         =   "Return to Main Menu"
      Height          =   855
      Left            =   240
      TabIndex        =   21
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Cmdbeauty 
      Caption         =   "Click for your score!"
      Height          =   975
      Left            =   240
      TabIndex        =   20
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Frame Frame5 
      Caption         =   "Who Loves Belle"
      Height          =   975
      Left            =   6120
      TabIndex        =   16
      Top             =   4440
      Width           =   4215
      Begin VB.OptionButton Option15 
         Caption         =   "Cogsworth"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Lumiere"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Gaston"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "4. What does Belle love to do?"
      Height          =   975
      Left            =   6120
      TabIndex        =   12
      Top             =   3360
      Width           =   4215
      Begin VB.OptionButton Option12 
         Caption         =   "Go Hunting"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Ride Horses"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2895
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Read"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "3. Who turned the beast into a beast?"
      Height          =   975
      Left            =   6120
      TabIndex        =   8
      Top             =   2280
      Width           =   4215
      Begin VB.OptionButton Option9 
         Caption         =   "Enchantress"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Fairy"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Gypsy"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "2. Why does Belle stay at the castle?"
      Height          =   975
      Left            =   6120
      TabIndex        =   4
      Top             =   1200
      Width           =   4215
      Begin VB.OptionButton Option6 
         Caption         =   "To save her father"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Because she likes the castle"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Because she likes the Beast"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "1. What was Belle's Horses name?"
      Height          =   1095
      Left            =   6120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton Option3 
         Caption         =   "Pierre"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Phillipe"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "George"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FrmForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ctr As Integer
'Movie Trivia
'FrmBeauty
'Amber Olson, Emily Borka, Shannon O'Neill
'11-1-08
' The purpose of this form is to allow the user to click their answer in Radio Buttons and then get a score, like a quiz.


Private Sub Cmdbeauty_Click()
'These are the answers to the radio buttons. Using boolean values to give the user a score after they are done with the quiz.
'And print them out.
Ctr = 0
If Option2 = True Then
   Ctr = Ctr + 1
End If

If Option6 = True Then
    Ctr = Ctr + 1
End If

If Option9 = True Then
    Ctr = Ctr + 1
End If

If Option10 = True Then
    Ctr = Ctr + 1
End If

If Option13 = True Then
    Ctr = Ctr + 1
End If

MsgBox UserName & " you got " & Ctr & " right out of 5.", , "Score"

End Sub

Private Sub Cmdmenu_Click()
'This is to get all of the radio buttons to clear and to get back to the main page with all of the other pages hidden.
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
Option12.Value = False
Option13.Value = False
Option14.Value = False
Option15.Value = False
FrmForm1.Show
FrmForm2.Hide
FrmForm3.Hide
FrmForm4.Hide
Frmsearch.Hide
End Sub
