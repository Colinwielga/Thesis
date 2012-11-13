VERSION 5.00
Begin VB.Form frmQuizFDR 
   Caption         =   "FDR Quiz"
   ClientHeight    =   7485
   ClientLeft      =   4065
   ClientTop       =   2160
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   7500
   Begin VB.PictureBox picAmericanQuizBC 
      Height          =   7695
      Left            =   -360
      Picture         =   "frmQuizFDR.frx":0000
      ScaleHeight     =   7635
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   -120
      Width           =   9135
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
         TabIndex        =   35
         Top             =   960
         Width           =   5295
         Begin VB.OptionButton opt1TrueFDR 
            Caption         =   "New York"
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
            TabIndex        =   41
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton optFalse1FDR 
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
            TabIndex        =   40
            Top             =   1920
            Width           =   855
         End
         Begin VB.OptionButton optFalse2FDR 
            Caption         =   "Hawaii"
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
            Width           =   1095
         End
         Begin VB.OptionButton optFalse3FDR 
            Caption         =   "Prussia"
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
            Width           =   1335
         End
         Begin VB.CommandButton cmdAnswerFDR1 
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
            TabIndex        =   36
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label lbl1 
            Alignment       =   2  'Center
            Caption         =   "In what state was President Roosevelt born in?"
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
            TabIndex        =   42
            Top             =   480
            Width           =   3615
         End
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
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         Begin VB.OptionButton opt2False1FDR 
            Caption         =   "5 term, 9 years"
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
            TabIndex        =   33
            Top             =   1560
            Width           =   1575
         End
         Begin VB.OptionButton opt2False2FDR 
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
            Height          =   225
            Left            =   720
            TabIndex        =   32
            Top             =   1920
            Width           =   1815
         End
         Begin VB.OptionButton opt2TrueFDR 
            Caption         =   "3 terms, 12 years"
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
            Top             =   2280
            Width           =   1815
         End
         Begin VB.OptionButton opt2False3FDR 
            Caption         =   "3 terms, 18 years"
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
            Top             =   2640
            Width           =   1815
         End
         Begin VB.CommandButton cmdAnswerFDR2 
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
            TabIndex        =   29
            Top             =   1440
            Width           =   2175
         End
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
            TabIndex        =   28
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label lbl2 
            Alignment       =   2  'Center
            Caption         =   "How long did FDR serve as president?"
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
         TabIndex        =   26
         Top             =   6960
         Width           =   1095
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
         TabIndex        =   25
         Top             =   6960
         Width           =   1095
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
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
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
            TabIndex        =   23
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdAnswerFDR3 
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
            TabIndex        =   22
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton opt3False3FDR 
            Caption         =   "Jackson Pollock, 1932-1949"
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
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Width           =   2655
         End
         Begin VB.OptionButton opt3False2FDR 
            Caption         =   "Jesse Jackson, 1930-1950"
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
            Left            =   120
            TabIndex        =   20
            Top             =   2280
            Width           =   2415
         End
         Begin VB.OptionButton opt3TrueFDR 
            Caption         =   "John Garner, 1932-1945"
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
            Left            =   120
            TabIndex        =   19
            Top             =   1920
            Width           =   2295
         End
         Begin VB.OptionButton opt3False1FDR 
            Caption         =   "Fran Tarkenton, 1935-1939"
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
            Left            =   120
            TabIndex        =   18
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label lbl3 
            Alignment       =   2  'Center
            Caption         =   "Who was Roosevelt's vice president for the longest period of time?"
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
            Left            =   360
            TabIndex        =   24
            Top             =   600
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
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         Begin VB.OptionButton opt4False1FDR 
            Caption         =   "1941"
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
            TabIndex        =   15
            Top             =   2640
            Width           =   735
         End
         Begin VB.OptionButton opt4False2FDR 
            Caption         =   "1947"
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
            Top             =   1920
            Width           =   735
         End
         Begin VB.OptionButton opt4False3FDR 
            Caption         =   "1904"
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
            Width           =   735
         End
         Begin VB.OptionButton opt4TrueFDR 
            Caption         =   "1932"
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
            Top             =   1560
            Width           =   735
         End
         Begin VB.CommandButton cmdAnswerFDR4 
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
            TabIndex        =   11
            Top             =   1440
            Width           =   2175
         End
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
            TabIndex        =   10
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label lbl4 
            Alignment       =   2  'Center
            Caption         =   "When was President Roosevelt first elected?"
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
            TabIndex        =   16
            Top             =   720
            Width           =   4575
         End
      End
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
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
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
            TabIndex        =   7
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdAnswerFDR5 
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
         Begin VB.OptionButton opt5False3FDR 
            Caption         =   "1st"
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
            Width           =   735
         End
         Begin VB.OptionButton opt5False2FDR 
            Caption         =   "107th"
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
            Top             =   2640
            Width           =   855
         End
         Begin VB.OptionButton opt5False1FDR 
            Caption         =   "13th"
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
            Width           =   735
         End
         Begin VB.OptionButton opt5TrueFDR 
            Caption         =   "32nd"
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
            TabIndex        =   2
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lbl5 
            Alignment       =   2  'Center
            Caption         =   "What number president was FDR?"
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
            TabIndex        =   8
            Top             =   720
            Width           =   4575
         End
      End
   End
End
Attribute VB_Name = "frmQuizFDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAnswerFDR1_Click()
    If opt1TrueFDR = True Then
        MsgBox ("Congratulations on getting the first problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext1.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizFDR.Hide
    End If
    
End Sub

Private Sub cmdAnswerFDR2_Click()
    If opt2TrueFDR = True Then
        MsgBox ("Congratulations on getting the second problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext2.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizFDR.Hide
    End If
    
End Sub
Private Sub cmdAnswerFDR3_Click()
    If opt3TrueFDR = True Then
        MsgBox ("Congratulations on getting the third problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext3.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizFDR.Hide
    End If
    
End Sub

Private Sub cmdAnswerFDR4_Click()
    If opt4TrueFDR = True Then
        MsgBox ("Congratulations on getting the fourth problem correct!")
        MsgBox ("Click Next when you are ready.")
        cmdNext4.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizFDR.Hide
    End If
    
End Sub

Private Sub cmdAnswerFDR5_Click()
    If opt5TrueFDR = True Then
        MsgBox ("Congratulations on getting the fifth and final problem correct!")
        MsgBox ("Click Grand Prize when you are ready to go Mad Lib President Roosevelt's 1932 Inauguration Speech.")
        cmdPrize.Enabled = True
    Else
        MsgBox ("Sorry, but it seems that you have the wrong answer selected.  You must now start from the main page.  Good Luck Next Time!")
        frmBeginMadLib.Show
        frmQuizFDR.Hide
    End If
    
End Sub

Private Sub cmdGoBack_Click()
    frmBeginMadLib.Show
    frmQuizFDR.Hide
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
    frmQuizFDR.Visible = False
    frmFDR.Visible = True
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

