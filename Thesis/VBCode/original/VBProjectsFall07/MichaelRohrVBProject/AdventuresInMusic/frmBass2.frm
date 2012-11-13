VERSION 5.00
Begin VB.Form frmBass2 
   Caption         =   "Learning the Bass Clef 2"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   Picture         =   "frmBass2.frx":0000
   ScaleHeight     =   8265
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinish 
      BackColor       =   &H0080C0FF&
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
      TabIndex        =   34
      Top             =   6960
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
      TabIndex        =   33
      Top             =   6960
      Width           =   1215
   End
   Begin VB.PictureBox picBass2 
      AutoSize        =   -1  'True
      Height          =   2220
      Left            =   1800
      Picture         =   "frmBass2.frx":32685A
      ScaleHeight     =   2160
      ScaleWidth      =   7335
      TabIndex        =   20
      Top             =   3600
      Width           =   7395
      Begin VB.Line Line6 
         X1              =   5040
         X2              =   5400
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line5 
         X1              =   3960
         X2              =   4320
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line4 
         X1              =   3240
         X2              =   3840
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line3 
         X1              =   2880
         X2              =   3840
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line2 
         X1              =   2160
         X2              =   2880
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   1560
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
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
         Left            =   6720
         TabIndex        =   30
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
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
         Left            =   6120
         TabIndex        =   29
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FF80FF&
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
         Left            =   5640
         TabIndex        =   28
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Left            =   5040
         TabIndex        =   27
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   4440
         TabIndex        =   26
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Left            =   3960
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   24
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   23
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   22
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
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
         Left            =   1680
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.TextBox txt1 
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
      Left            =   2040
      TabIndex        =   9
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txt2 
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
      Left            =   3720
      TabIndex        =   8
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txt3 
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
      Left            =   5400
      TabIndex        =   7
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txt4 
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
      Left            =   7080
      TabIndex        =   6
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txt5 
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
      Left            =   8640
      TabIndex        =   5
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txt6 
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
      Left            =   2040
      TabIndex        =   4
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txt7 
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
      Left            =   3720
      TabIndex        =   3
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txt8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
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
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txt9 
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
      Left            =   7080
      TabIndex        =   1
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txt10 
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
      Left            =   8640
      TabIndex        =   0
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label lblBass2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   $"frmBass2.frx":35A25C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1695
      Left            =   840
      TabIndex        =   32
      Top             =   1800
      Width           =   9135
   End
   Begin VB.Label lblBass0 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Notes on the Bass Clef"
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   1080
      TabIndex        =   31
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
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
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
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
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
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
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
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
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
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
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
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
      Top             =   6960
      Width           =   375
   End
End
Attribute VB_Name = "frmBass2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is meant to quiz the user about there knowledge of the bass clef asking 10 questions
'the form uses textboxes for the user to input the answers into when the user puts in a
'correct answer there score is incremented by 1 point in the Public variable BassPoints

Private Sub cmdBack_Click()         'This button changes forms
    frmBass2.Hide                   'this hides frmBass2
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
End Sub

Private Sub cmdFinish_Click()       'This button takes the information given in the textboxes,
                                    'sees if they are correct using if statements, and if so the variable BassPoints is incremented by 1
    BassPoints = 0                      'sets the value of the Public variable BassPoints = 0
    If txt1.Text = "f" Then             'If statement checks the correct answer for txt1.Text
        BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    If txt2.Text = "g" Then             'If statement checks the correct answer for txt2.Text
        BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    If txt3.Text = "b" Then             'If statement checks the correct answer for txt3.Text
        BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    If txt4.Text = "c" Then             'If statement checks the correct answer for txt4.Text
        BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    If txt5.Text = "a" Then             'If statement checks the correct answer for txt5.Text
        BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    If txt6.Text = "e" Then             'If statement checks the correct answer for txt6.Text
        BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    If txt7.Text = "d" Then             'If statement checks the correct answer for txt7.Text
    BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    If txt8.Text = "c" Then             'If statement checks the correct answer for txt8.Text
        BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    If txt9.Text = "g" Then             'If statement checks the correct answer for txt9.Text
        BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    If txt10.Text = "a" Then            'If statement checks the correct answer for txt10.Text
        BassPoints = BassPoints + 1     'if correct adds 1 to value of BassPoints
    End If
    MsgBox "Congratulations " & NameGiven & "!!!  You got " & BassPoints & " points!!!", , "Your Score"     'displays a message box giving the user, with NameGiven Public variable,
                                                                                                            'there score, BassPoints, from the quiz
End Sub
