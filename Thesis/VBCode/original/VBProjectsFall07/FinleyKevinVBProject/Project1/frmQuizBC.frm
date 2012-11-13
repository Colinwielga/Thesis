VERSION 5.00
Begin VB.Form frmQuizBC 
   Caption         =   "Bill Clinton Quiz"
   ClientHeight    =   7500
   ClientLeft      =   3960
   ClientTop       =   2160
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7500
   Begin VB.PictureBox picAmericanQuizBC 
      Height          =   7695
      Left            =   -360
      Picture         =   "frmQuizBC.frx":0000
      ScaleHeight     =   7635
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   -120
      Width           =   9135
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
         Begin VB.OptionButton opt5TrueBC 
            Caption         =   "42nd"
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
         Begin VB.OptionButton opt5False1BC 
            Caption         =   "12th"
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
         Begin VB.OptionButton opt5False2BC 
            Caption         =   "90th"
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
            Top             =   2280
            Width           =   735
         End
         Begin VB.OptionButton opt5False3BC 
            Caption         =   "40th"
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
            Top             =   2640
            Width           =   735
         End
         Begin VB.CommandButton cmdAnswerBC5 
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "What number president was Bill Clinton?"
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
         Begin VB.CommandButton cmdAnswerBC4 
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
         Begin VB.OptionButton opt4TrueBC 
            Caption         =   "1992"
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
            Top             =   2640
            Width           =   735
         End
         Begin VB.OptionButton opt4False3BC 
            Caption         =   "1998"
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
         Begin VB.OptionButton opt4False2BC 
            Caption         =   "1994"
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
            Top             =   1920
            Width           =   735
         End
         Begin VB.OptionButton opt4False1BC 
            Caption         =   "1937"
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
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label lbl4 
            Alignment       =   2  'Center
            Caption         =   "When was President Clinton first elected?"
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
            Top             =   720
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
         Begin VB.OptionButton opt3False1BC 
            Caption         =   "Adrian Peterson"
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
            TabIndex        =   25
            Top             =   1560
            Width           =   1695
         End
         Begin VB.OptionButton opt3TrueBC 
            Caption         =   "Al Gore"
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
            TabIndex        =   24
            Top             =   1920
            Width           =   975
         End
         Begin VB.OptionButton opt3False2BC 
            Caption         =   "Bill Richardson"
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
            TabIndex        =   23
            Top             =   2280
            Width           =   1935
         End
         Begin VB.OptionButton opt3False3BC 
            Caption         =   "Mike Gravel"
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
            TabIndex        =   22
            Top             =   2640
            Width           =   2415
         End
         Begin VB.CommandButton cmdAnswerBC3 
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
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Who was Bill Clinton's vice president?"
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
            TabIndex        =   26
            Top             =   720
            Width           =   4575
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
         Left            =   5280
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
         Left            =   6600
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
            TabIndex        =   16
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdAnswerBC2 
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
            TabIndex        =   15
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton opt2False3BC 
            Caption         =   "4 terms, 20 years"
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
            TabIndex        =   14
            Top             =   2640
            Width           =   2415
         End
         Begin VB.OptionButton opt2TrueBC 
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
            TabIndex        =   13
            Top             =   2280
            Width           =   1935
         End
         Begin VB.OptionButton opt2False2BC 
            Caption         =   "2 terms, 12 years"
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
            TabIndex        =   12
            Top             =   1920
            Width           =   2055
         End
         Begin VB.OptionButton opt2False1BC 
            Caption         =   "1 term, 2 years"
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
            TabIndex        =   11
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "How long did Bill Clinton serve as president?"
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
            TabIndex        =   10
            Top             =   720
            Width           =   4575
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
            TabIndex        =   8
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdAnswerBC1 
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
            TabIndex        =   7
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton optFalse3BC 
            Caption         =   "Afghanistan"
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
            TabIndex        =   6
            Top             =   2640
            Width           =   1335
         End
         Begin VB.OptionButton optFalse2BC 
            Caption         =   "Alabama"
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
            Top             =   2280
            Width           =   1095
         End
         Begin VB.OptionButton optFalse1BC 
            Caption         =   "Alaska"
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
            Top             =   1920
            Width           =   975
         End
         Begin VB.OptionButton optTrue1BC 
            Caption         =   "Arkansas"
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
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "In what state was President Bill Clinton born in?"
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
            TabIndex        =   2
            Top             =   720
            Width           =   3615
         End
      End
   End
End
Attribute VB_Name = "frmQuizBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAnswerBC1_Click()
    If optTrue1BC = True Then
        MsgBox ("Congratulations on getting the first problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext1.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizBC.Hide
    End If
End Sub

Private Sub cmdAnswerBC2_Click()
    If opt2TrueBC = True Then
        MsgBox ("Congratulations on getting the second problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext2.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizBC.Hide
    End If
End Sub

Private Sub cmdAnswerBC3_Click()
    If opt3TrueBC = True Then
        MsgBox ("Congratulations on getting the third problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext3.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizBC.Hide
    End If
End Sub

Private Sub cmdAnswerBC4_Click()
    If opt4TrueBC = True Then
        MsgBox ("Congratulations on getting the fourth problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext4.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizBC.Hide
    End If
End Sub

Private Sub cmdAnswerBC5_Click()
    If opt5TrueBC = True Then
        MsgBox ("Congratulations on getting the fifth and final problem correct!")
        MsgBox ("Click Grand Prize when you are ready to go Mad Lib on President Clinton's 1997 Inauguration Speech.")
        cmdPrize.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizBC.Hide
    End If
End Sub

Private Sub cmdGoBack_Click()
    frmBeginMadLib.Show
    frmQuizBC.Hide
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
    frmQuizBC.Visible = False
    frmBC.Visible = True
    
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

