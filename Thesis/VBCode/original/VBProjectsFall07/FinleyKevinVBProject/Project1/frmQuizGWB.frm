VERSION 5.00
Begin VB.Form frmQuizGWB 
   Caption         =   "George Bush Quiz"
   ClientHeight    =   7530
   ClientLeft      =   3960
   ClientTop       =   2160
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   7545
   Begin VB.PictureBox picAmericanQuizBC 
      Height          =   7695
      Left            =   -360
      Picture         =   "frmQuizGWB.frx":0000
      ScaleHeight     =   7635
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   -120
      Width           =   7935
      Begin VB.Frame frame5 
         Caption         =   "FINAL Question #5"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   1440
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         Begin VB.OptionButton opt5TrueGWB 
            Caption         =   "43rd"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   41
            Top             =   1560
            Width           =   735
         End
         Begin VB.OptionButton opt5False1GWB 
            Caption         =   "99th"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   40
            Top             =   1920
            Width           =   735
         End
         Begin VB.OptionButton opt5False2GWB 
            Caption         =   "45th"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   39
            Top             =   2640
            Width           =   855
         End
         Begin VB.OptionButton opt5False3GWB 
            Caption         =   "22nd"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   38
            Top             =   2280
            Width           =   735
         End
         Begin VB.CommandButton cmdAnswerGWB5 
            Caption         =   "When you think you have the right answer, CLICK this button."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            TabIndex        =   37
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CommandButton cmdPrize 
            Caption         =   "Grand Prize!"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            TabIndex        =   36
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label lbl5 
            Alignment       =   2  'Center
            Caption         =   "What number president is George Bush?"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   42
            Top             =   720
            Width           =   4575
         End
      End
      Begin VB.Frame frame4 
         Caption         =   "Question #4"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   1440
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         Begin VB.CommandButton cmdNext4 
            Caption         =   "Next"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            TabIndex        =   33
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdAnswerGWB4 
            Caption         =   "When you think you have the right answer, CLICK this button."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            TabIndex        =   32
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton opt4TrueGWB 
            Caption         =   "2000"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   31
            Top             =   1920
            Width           =   735
         End
         Begin VB.OptionButton opt4False3GWB 
            Caption         =   "1776"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   30
            Top             =   2280
            Width           =   735
         End
         Begin VB.OptionButton opt4False2GWB 
            Caption         =   "1978"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   29
            Top             =   1560
            Width           =   735
         End
         Begin VB.OptionButton opt4False1GWB 
            Caption         =   "1996"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   28
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label lbl4 
            Alignment       =   2  'Center
            Caption         =   "When was President Bush first elected?"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   34
            Top             =   840
            Width           =   4575
         End
      End
      Begin VB.Frame frame3 
         Caption         =   "Question #3"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   1440
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         Begin VB.OptionButton opt3False1GWB 
            Caption         =   "Alan Page"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   25
            Top             =   2640
            Width           =   1215
         End
         Begin VB.OptionButton opt3TrueGWB 
            Caption         =   "Dick Cheney"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   24
            Top             =   1920
            Width           =   1455
         End
         Begin VB.OptionButton opt3False2GWB 
            Caption         =   "Paul Revere"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   2280
            Width           =   1335
         End
         Begin VB.OptionButton opt3False3GWB 
            Caption         =   "Mitt Romney"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   22
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdAnswerGWB3 
            Caption         =   "When you think you have the right answer, CLICK this button."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            TabIndex        =   21
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CommandButton cmdNext3 
            Caption         =   "Next"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            TabIndex        =   20
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label lbl3 
            Alignment       =   2  'Center
            Caption         =   "Who is Bush's vice president?"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   26
            Top             =   720
            Width           =   3735
         End
      End
      Begin VB.CommandButton cmdGoBack 
         Caption         =   "Go Back"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   18
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   17
         Top             =   6960
         Width           =   1095
      End
      Begin VB.Frame frame2 
         Caption         =   "Question #2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         Begin VB.CommandButton cmdNext2 
            Caption         =   "Next"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            TabIndex        =   15
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdAnswerGWB2 
            Caption         =   "When you think you have the right answer, CLICK this button."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            TabIndex        =   14
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton opt2False3GWB 
            Caption         =   "4 terms, 4 years"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   13
            Top             =   2640
            Width           =   1815
         End
         Begin VB.OptionButton opt2TrueGWB 
            Caption         =   "2 terms, 8 years"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   12
            Top             =   2280
            Width           =   1815
         End
         Begin VB.OptionButton opt2False2GWB 
            Caption         =   "1 term, 15 years"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   720
            TabIndex        =   11
            Top             =   1920
            Width           =   1815
         End
         Begin VB.OptionButton opt2False1GWB 
            Caption         =   "3 terms, 9 years"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   10
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lbl2 
            Alignment       =   2  'Center
            Caption         =   "How long will Bush have served when his current term is up?"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   480
            TabIndex        =   16
            Top             =   720
            Width           =   4455
         End
      End
      Begin VB.Frame frame1 
         Caption         =   "Question #1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   5295
         Begin VB.CommandButton cmdNext1 
            Caption         =   "Next"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            TabIndex        =   7
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdAnswerGWB1 
            Caption         =   "When you think you have the right answer, CLICK this button."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            TabIndex        =   6
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton optFalse3GWB 
            Caption         =   "Montana"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   5
            Top             =   1560
            Width           =   1335
         End
         Begin VB.OptionButton optFalse2GWB 
            Caption         =   "Vermont"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   4
            Top             =   2280
            Width           =   1095
         End
         Begin VB.OptionButton optFalse1GWB 
            Caption         =   "Texas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   3
            Top             =   1920
            Width           =   855
         End
         Begin VB.OptionButton opt1TrueGWB 
            Caption         =   "Connecticut"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   2
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lbl1 
            Alignment       =   2  'Center
            Caption         =   "In what state was President Bush born in?"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   840
            TabIndex        =   8
            Top             =   480
            Width           =   3615
         End
      End
   End
End
Attribute VB_Name = "frmQuizGWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAnswerGWB1_Click()
    If opt1TrueGWB = True Then
        MsgBox ("Congratulations on getting the first problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext1.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizGWB.Hide
    End If
End Sub
Private Sub cmdAnswerGWB2_Click()
    If opt2TrueGWB = True Then
        MsgBox ("Congratulations on getting the second problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext2.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizGWB.Hide
    End If
End Sub
Private Sub cmdAnswerGWB3_Click()
    If opt3TrueGWB = True Then
        MsgBox ("Congratulations on getting the third problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext3.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizGWB.Hide
    End If
End Sub
Private Sub cmdAnswerGWB4_Click()
    If opt4TrueGWB = True Then
        MsgBox ("Congratulations on getting the fourth problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext4.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizGWB.Hide
    End If
End Sub
Private Sub cmdAnswerGWB5_Click()
    If opt5TrueGWB = True Then
        MsgBox ("Congratulations on getting the fifth and final problem correct!")
        MsgBox ("Click Next when you are ready to go Mad Lib President Bush's 2000 Inauguration Speech.")
        cmdPrize.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizGWB.Hide
    End If
End Sub

Private Sub cmdGoBack_Click()
    frmBeginMadLib.Show
    frmQuizGWB.Hide
End Sub

Private Sub cmdNext1_Click()
    frame1.Visible = False
    frame2.Visible = True
End Sub
Private Sub cmdNext2_Click()
    frame2.Visible = False
    frame3.Visible = True
End Sub

Private Sub cmdNext3_Click()
    frame3.Visible = False
    frame4.Visible = True
End Sub

Private Sub cmdNext4_Click()
    frame4.Visible = False
    frame5.Visible = True
End Sub

Private Sub cmdPrize_Click()
    frmQuizGWB.Visible = False
    frmGWB.Visible = True
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

