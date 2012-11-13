VERSION 5.00
Begin VB.Form frmTreble2 
   Caption         =   "Learning the Treble Clef 2"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   Picture         =   "frmTreble2.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTreble 
      Height          =   2925
      Left            =   1560
      Picture         =   "frmTreble2.frx":2BBACA
      ScaleHeight     =   2865
      ScaleWidth      =   7935
      TabIndex        =   24
      Top             =   3360
      Width           =   7995
      Begin VB.Line Line4 
         X1              =   4560
         X2              =   4920
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line5 
         X1              =   5760
         X2              =   6120
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "9."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   34
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "8."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   33
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Caption         =   "7."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   32
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "6."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   31
         Top             =   1320
         Width           =   375
      End
      Begin VB.Line Line3 
         X1              =   3960
         X2              =   4560
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "5."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   30
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "4."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   29
         Top             =   720
         Width           =   375
      End
      Begin VB.Line Line2 
         X1              =   2760
         X2              =   3360
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   28
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   27
         Top             =   1440
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   2280
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "10."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   25
         Top             =   2280
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdFinish 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Finished"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back To Main Page"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txt4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txt5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   5
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txt6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox txt7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox txt8 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox txt9 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox txt10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   0
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblBass2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   $"frmTreble2.frx":30BF18
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1695
      Left            =   720
      TabIndex        =   21
      Top             =   1440
      Width           =   9135
   End
   Begin VB.Label lblTreble 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Notes on theTreble Clef"
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1215
      Left            =   960
      TabIndex        =   20
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "6."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "7."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "8."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "9."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Top             =   7320
      Width           =   375
   End
End
Attribute VB_Name = "frmTreble2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is set up to be a quiz for the user to test their knowledge about the Treble clef and the notes upon the staff
'It uses text boxes like the quiz on the piano keyboard and also the bass clef quiz
'The user inputs a letter in lower case and when finished clicks the finished button and a bunch of If statements evaluate whether or not the question is correct
'if it is, the user gets + 1 points, if not, nothing happens to the score

Private Sub cmdBack_Click()     'This button changes forms to frmLessonMainPage
    frmTreble2.Hide                 'this hides frmTreble2
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
End Sub

'This button evaluates the answers from the many text boxes on the form
'if the answer is the correct letter (and in lower case) then the user is given + 1 points, if not, nothing happens to the score
'at the end a message box pops up and tells the user their score
Private Sub cmdFinish_Click()
    TreblePoints = 0
    If txt1.Text = "d" Then                 'If statement checks the correct answer for txt1.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    If txt2.Text = "a" Then                 'If statement checks the correct answer for txt2.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    If txt3.Text = "g" Then                 'If statement checks the correct answer for txt3.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    If txt4.Text = "e" Then                 'If statement checks the correct answer for txt4.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    If txt5.Text = "f" Then                 'If statement checks the correct answer for txt5.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    If txt6.Text = "b" Then                 'If statement checks the correct answer for txt6.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    If txt7.Text = "c" Then                 'If statement checks the correct answer for txt7.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    If txt8.Text = "e" Then                 'If statement checks the correct answer for txt8.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    If txt9.Text = "f" Then                 'If statement checks the correct answer for txt9.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    If txt10.Text = "c" Then                'If statement checks the correct answer for txt10.Text
        TreblePoints = TreblePoints + 1     'if correct adds 1 to value of TreblePoints
    End If
    MsgBox "Congratulations " & NameGiven & "!!!  You got " & TreblePoints & " points!!!", , "Your Score"       'Giving a personal touch to the message box, the user's NameGiven is used and their score is told using the variable Public TreblePoints
End Sub
