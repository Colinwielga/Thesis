VERSION 5.00
Begin VB.Form frmTakeQuiz2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Take Quiz"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FF80FF&
      Caption         =   "View Your Celeb Style!"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   3615
      Begin VB.OptionButton Opt9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Opt9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Opt9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Opt9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3735
      Begin VB.OptionButton Opt7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   0
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Opt7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton Opt7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Opt7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00FF80FF&
      Caption         =   "Submit Results"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2055
      Left            =   5040
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
      Begin VB.OptionButton Opt8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Opt8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Opt8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Opt8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Height          =   2535
      Left            =   4920
      TabIndex        =   8
      Top             =   5520
      Width           =   3015
      Begin VB.OptionButton Opt10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   24
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton Opt10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   0
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Opt10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Opt10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label lblNames 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenna Gebeke ~ Katie Ranallo"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   26
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label lblQuestion10 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz2.frx":0000
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4920
      TabIndex        =   3
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblQuestion9 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz2.frx":00A8
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label lblQuestion8 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz2.frx":01F7
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblQuestion7 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz2.frx":02A4
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmTakeQuiz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Take Quiz 2 (questions 7-10)
'Form Objective: 'Form Objective: This allows the user to read quiz questions 7-10 and select answers(A, B, C, or D) that indicate their response. This form also allows the user to submit their results, which will then display their results in a message box, which will direct them to select the command button- View Your Celeb Match from this form.
Private Sub cmdSubmit_Click()
'This command button allows the user to submit their answers once they have completed the quiz after the answers are counted.  The user's results will be displayed in a message box, which will direct them to select the View Your Celeb Match command button.
    Dim Max As String
    
    If Opt7(1).Value = True Then
        Actr = Actr + 1
    End If
    If Opt7(2).Value = True Then
        Bctr = Bctr + 1
    End If
    If Opt7(3).Value = True Then
        Cctr = Cctr + 1
    End If
    If Opt7(4).Value = True Then
        Dctr = Dctr + 1
    End If
    If Opt8(1).Value = True Then
        Actr = Actr + 1
    End If
    If Opt8(2).Value = True Then
        Bctr = Bctr + 1
    End If
    If Opt8(3).Value = True Then
        Cctr = Cctr + 1
    End If
    If Opt8(4).Value = True Then
        Dctr = Dctr + 1
    End If
    If Opt9(1).Value = True Then
        Actr = Actr + 1
    End If
    If Opt9(2).Value = True Then
        Bctr = Bctr + 1
    End If
    If Opt9(3).Value = True Then
        Cctr = Cctr + 1
    End If
    If Opt9(4).Value = True Then
        Dctr = Dctr + 1
    End If
    If Opt10(1).Value = True Then
        Actr = Actr + 1
    End If
    If Opt10(2).Value = True Then
        Bctr = Bctr + 1
    End If
    If Opt10(3).Value = True Then
        Cctr = Cctr + 1
    End If
    If Opt10(4).Value = True Then
        Dctr = Dctr + 1
    End If
    
    'These nested If-Then statements compare each counter (the number of As,Bs,Cs,Ds selected by the user) to find the maximum counter. The maximum counter is the most freqently selected answer.
    'When the maximum counter is determined,a message box will be displayed with the user's results.
    'The results of the quiz are then inputed into a file that holds quiz takers.  It inputs their user name and their style.
    If Actr > Bctr Then
        If Actr > Cctr Then
            If Actr > Dcrt Then
                Max = A
                MsgBox "Your style is Spontaneously Sporty! Your clothes reflect your active lifestyle.  You relish and feel most comfortable in stretchy cotton fabrics that allow you to be spontaneous and ready for action.  Your celeb style match is Jennifer Garner. To read more about Jennifer's style, click View Your Celeb Match.", , "Your Style is Spontaneously Sporty!"
                Open App.Path & "\QuizTakers.txt" For Append As #1
                Write #1, userName, "Spontaneously Sporty!"
                Close #1
            End If
        End If
    
    ElseIf Bctr > Cctr Then
        If Bctr > Dctr Then
            If Bctr > Acrt Then
                Max = B
                MsgBox "Your style is Classically Chic! Looking pulled together at all times is a priority for you.  That's why you rely on classic, simple items that don't require a lot of thought and have timeless style.  Your celeb style match is Reese Witherspoon. To read more about Reese's style, click View Your Celeb Match.", , "Your Style is Classically Chic!"
                Open App.Path & "\QuizTakers.txt" For Append As #1
                Write #1, userName, "Classically Chic!"
                Close #1
            End If
        End If
        
     ElseIf Cctr > Dctr Then
        If Cctr > Actr Then
            If Cctr > Bcrt Then
                Max = C
                MsgBox "Your style is Subtly Sexy! Your clothes reflect your sexy nature.  You relish the feel of luxurious fabrics against your skin, and you dress to bring out your alluring side.  Your celeb style match is Jennifer Lopez. To read more about JLo's style, click View Your Celeb Match.", , "Your Style is Subtly Sexy!"
                Open App.Path & "\QuizTakers.txt" For Append As #1
                Write #1, userName, "Subtly Sexy!"
                Close #1
            End If
        End If
        
    ElseIf Dctr > Actr Then
        If Dctr > Bctr Then
            If Dctr > Ccrt Then
                Max = D
                MsgBox "Your style is Tragically Trendy! You try so hard to show off the latest styles that sometimes it's a miss but when you're on, it's a hit! You focus on loud, gaudy items that are only in the here and now.  Your celeb style match is Paris Hilton.  To read more about Paris's style, click View Your Celeb Match.  ", , "Your Style is Tragically Trendy!"
                Open App.Path & "\QuizTakers.txt" For Append As #1
                Write #1, userName, "Tragically Trendy!"
                Close #1
            End If
        End If
    End If
    
End Sub

Private Sub cmdView_Click()
'This command button allows the user to view their celeb style match found on the Celebs form.
    frmTakeQuiz2.Hide
    frmCelebs.Show
End Sub

Private Sub Form_Load()
    frmTakeQuiz2.Caption = "Welcome " & userName & "  - Take Quiz"
End Sub
