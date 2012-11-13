VERSION 5.00
Begin VB.Form FrmForm3 
   BackColor       =   &H00FF8080&
   Caption         =   "Back to the Future"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14100
   LinkTopic       =   "Form3"
   Picture         =   "FrmBack.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to the Menu"
      Height          =   735
      Left            =   10920
      TabIndex        =   21
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdTotalScore 
      Caption         =   "Total Your Score"
      Height          =   735
      Left            =   10920
      TabIndex        =   20
      Top             =   360
      Width           =   2775
   End
   Begin VB.Frame Frame5 
      Caption         =   "What song does Marty play at the school dance?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4920
      TabIndex        =   16
      Top             =   5280
      Width           =   5535
      Begin VB.OptionButton Option15 
         Caption         =   """Peggy Sue"" by Buddy Holly"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   3855
      End
      Begin VB.OptionButton Option14 
         Caption         =   """Blitzkrieg Bop"" by The Ramones"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   3975
      End
      Begin VB.OptionButton Option13 
         Caption         =   """Johnny B. Goode"" by Chuck Berry"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "What is Doc Brown's signature catch phrase?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4920
      TabIndex        =   12
      Top             =   3960
      Width           =   5535
      Begin VB.OptionButton Option12 
         Caption         =   """Save the Cheerleader, Save the World."""
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   4455
      End
      Begin VB.OptionButton Option11 
         Caption         =   """Great Scott!"""
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   4455
      End
      Begin VB.OptionButton Option10 
         Caption         =   """That's Heavy."""
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "How much energy does it take to power the Flux Capacitor?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4920
      TabIndex        =   8
      Top             =   2640
      Width           =   5535
      Begin VB.OptionButton Option9 
         Caption         =   "1.21 Gigawatts"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Option8 
         Caption         =   "12.1 Megawatts"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option7 
         Caption         =   "121 Kilowatts"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "What was the directors' original design for the time machine?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4920
      TabIndex        =   4
      Top             =   1320
      Width           =   5535
      Begin VB.OptionButton Option6 
         Caption         =   "Electric Chair"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Refridgerator"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Washing Machine"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "What kind of car is the time machine made out of?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.OptionButton Option3 
         Caption         =   "Nissan Versa"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "DeLorean"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Dodge Viper"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Movie Trivia
'FrmBack
'Amber Olson, Emily Borka, Shannon O'Neill
'11-1-08
' The purpose of this form is to allow the user to click their answer in Radio Buttons and then get a score, like a quiz.

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

Private Sub cmdTotalScore_Click()
'These are the answers to the radio buttons. Using boolean values to give the user a score after they are done with the quiz.
'And print them out.
Dim Ctr As Integer
Ctr = 0
If Option2 = True Then
   Ctr = Ctr + 1
End If

If Option5 = True Then
    Ctr = Ctr + 1
End If

If Option9 = True Then
    Ctr = Ctr + 1
End If

If Option11 = True Then
    Ctr = Ctr + 1
End If

If Option13 = True Then
    Ctr = Ctr + 1
End If

MsgBox UserName & " you got " & Ctr & " right out of 5.", , "Score"

End Sub


