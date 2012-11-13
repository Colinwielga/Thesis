VERSION 5.00
Begin VB.Form frmTakeQuiz1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Take Quiz"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Height          =   2655
      Left            =   3720
      TabIndex        =   27
      Top             =   5640
      Width           =   3255
      Begin VB.OptionButton Opt5 
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
         TabIndex        =   31
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton Opt5 
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
         TabIndex        =   30
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton Opt5 
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
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Opt5 
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
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Height          =   2655
      Left            =   120
      TabIndex        =   22
      Top             =   5640
      Width           =   3255
      Begin VB.OptionButton Opt4 
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
         TabIndex        =   26
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton Opt4 
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
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton Opt4 
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
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Opt4 
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
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   2055
      Left            =   7200
      TabIndex        =   17
      Top             =   1680
      Width           =   2895
      Begin VB.OptionButton Opt3 
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
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton Opt3 
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
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Opt3 
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
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Opt3 
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
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1935
      Left            =   3720
      TabIndex        =   11
      Top             =   1800
      Width           =   2895
      Begin VB.OptionButton Opt2 
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
         TabIndex        =   16
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton Opt2 
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
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Opt2 
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
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Opt2 
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
         TabIndex        =   12
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
      Begin VB.OptionButton Opt1 
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
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton Opt1 
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
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton Opt1 
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
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Opt1 
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
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00FF80FF&
      Caption         =   "Continue Quiz..."
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      Height          =   1935
      Left            =   7200
      TabIndex        =   32
      Top             =   6360
      Width           =   2895
      Begin VB.OptionButton Opt6 
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
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Opt6 
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
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Opt6 
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
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Opt6 
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1095
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
      Height          =   375
      Left            =   7560
      TabIndex        =   37
      Top             =   8280
      Width           =   3015
   End
   Begin VB.Label lblQuestion5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz1.frx":0000
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3840
      TabIndex        =   6
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label lblQuestion4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz1.frx":0118
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label lblQuestion6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz1.frx":0259
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7200
      TabIndex        =   3
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label lblQuestion3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz1.frx":03D7
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7200
      TabIndex        =   2
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lblQuestion2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz1.frx":04D0
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblQuestion1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTakeQuiz1.frx":05E6
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmTakeQuiz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Take Quiz 1 (test questions 1-6)
'Form Objective: This allows the user to read quiz questions 1-6 and select answers(A, B, C, or D) that indicate their response. This form also includes a command button that allows the user to navigate to quiz questions 7-10.

Private Sub cmdContinue_Click()
'This command button allows the user to navigate to quiz questions 7-10 and increments the counters for the user's selections.
Actr = 0
Bctr = 0
Cctr = 0
Dctr = 0
'This determines if a value within the array of options is true or false. When an option is selected within the array, the value becomes true, and the counter of that option is incremented.
If Opt1(1).Value = True Then
    Actr = Actr + 1
End If
If Opt1(2).Value = True Then
    Bctr = Bctr + 1
End If
If Opt1(3).Value = True Then
    Cctr = Cctr + 1
End If
If Opt1(4).Value = True Then
    Dctr = Dctr + 1
End If
If Opt2(1).Value = True Then
    Actr = Actr + 1
End If
If Opt2(2).Value = True Then
    Bctr = Bctr + 1
End If
If Opt2(3).Value = True Then
    Cctr = Cctr + 1
End If
If Opt2(4).Value = True Then
    Dctr = Dctr + 1
End If
If Opt3(1).Value = True Then
    Actr = Actr + 1
End If
If Opt3(2).Value = True Then
    Bctr = Bctr + 1
End If
If Opt3(3).Value = True Then
    Cctr = Cctr + 1
End If
If Opt3(4).Value = True Then
    Dctr = Dctr + 1
End If
If Opt4(1).Value = True Then
    Actr = Actr + 1
End If
If Opt4(2).Value = True Then
    Bctr = Bctr + 1
End If
If Opt4(3).Value = True Then
    Cctr = Cctr + 1
End If
If Opt4(4).Value = True Then
    Dctr = Dctr + 1
End If
If Opt5(1).Value = True Then
    Actr = Actr + 1
End If
If Opt5(2).Value = True Then
    Bctr = Bctr + 1
End If
If Opt5(3).Value = True Then
    Cctr = Cctr + 1
End If
If Opt5(4).Value = True Then
    Dctr = Dctr + 1
End If
If Opt6(1).Value = True Then
    Actr = Actr + 1
End If
If Opt6(2).Value = True Then
    Bctr = Bctr + 1
End If
If Opt6(3).Value = True Then
    Cctr = Cctr + 1
End If
If Opt6(4).Value = True Then
    Dctr = Dctr + 1
End If
'The user is then directed to quiz questions 7-10.
    frmTakeQuiz1.Hide
    frmTakeQuiz2.Show
End Sub

Private Sub Form_Load()
    frmTakeQuiz1.Caption = "Welcome " & userName & "  - Take Quiz"
End Sub
