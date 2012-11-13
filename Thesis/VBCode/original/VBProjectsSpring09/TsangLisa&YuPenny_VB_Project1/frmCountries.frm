VERSION 5.00
Begin VB.Form frmCountries 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question: Which countries of cuisine do you want to make?"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCountries.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   10215
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAmerican 
      BackColor       =   &H008080FF&
      Caption         =   "American"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdMexican 
      BackColor       =   &H008080FF&
      Caption         =   "Mexican"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdChinese 
      BackColor       =   &H008080FF&
      Caption         =   "Chinese"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdItalian 
      BackColor       =   &H008080FF&
      Caption         =   "Italian"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   900
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton cmdFrench 
      BackColor       =   &H008080FF&
      Caption         =   "French"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton cmdJapanese 
      BackColor       =   &H008080FF&
      Caption         =   "Japanese"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   8760
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9360
      Width           =   975
   End
   Begin VB.Image img4 
      Height          =   3420
      Left            =   705
      Picture         =   "frmCountries.frx":08CA
      Top             =   5280
      Width           =   2745
   End
   Begin VB.Image img2 
      Height          =   3540
      Left            =   3960
      Picture         =   "frmCountries.frx":3703
      Top             =   1320
      Width           =   2610
   End
   Begin VB.Image img1 
      Height          =   3750
      Left            =   720
      Picture         =   "frmCountries.frx":5B68
      Top             =   1320
      Width           =   2550
   End
   Begin VB.Image img6 
      Height          =   3255
      Left            =   7185
      Picture         =   "frmCountries.frx":883B
      Top             =   5400
      Width           =   2685
   End
   Begin VB.Image img3 
      Height          =   3750
      Left            =   7320
      Picture         =   "frmCountries.frx":AC20
      Top             =   1320
      Width           =   2310
   End
   Begin VB.Image img5 
      Height          =   3240
      Left            =   3705
      Picture         =   "frmCountries.frx":E576
      Top             =   5280
      Width           =   2985
   End
   Begin VB.Image Image2 
      Height          =   2775
      Left            =   3600
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Which countries of cuisine do you want to make?"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   840
      TabIndex        =   7
      Top             =   360
      Width           =   9135
   End
   Begin VB.Shape shpQuestion 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00FFC0C0&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   660
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   9495
   End
End
Attribute VB_Name = "frmCountries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAmerican_Click(Index As Integer)

veggie = InputBox("Type in Yes if you are, No if you are not.", "Are you a vegetarian?")

If veggie = "Yes" Then
        frmCountries.Hide
        frmAmericanVeggie.Show
    ElseIf veggie = "No" Then
        frmCountries.Hide
        frmAmerican.Show
    Else: MsgBox ("Sorry, Please enter Yes or No.")
End If


End Sub

Private Sub cmdChinese_Click(Index As Integer)

veggie = InputBox("Type in Yes if you are, No if you are not.", "Are you a vegetarian?")

Select Case veggie
        Case Is = "Yes"
        frmCountries.Hide
        frmChineseVeggie.Show
Case "No"
        frmCountries.Hide
        frmChinese.Show
Case Else
        MsgBox ("Sorry, Please enter Yes or No.")
End Select


End Sub

Private Sub cmdFrench_Click(Index As Integer)

veggie = InputBox("Type in Yes if you are, No if you are not.", "Are you a vegetarian?")

If veggie = "Yes" Then
        frmCountries.Hide
        frmFrenchVeggie.Show
    ElseIf veggie = "No" Then
        frmCountries.Hide
        frmFrench.Show
    Else: MsgBox ("Sorry, Please enter Yes or No.")
End If

End Sub

Private Sub cmdItalian_Click(Index As Integer)
veggie = InputBox("Type in Yes if you are, No if you are not.", "Are you a vegetarian?")

If veggie = "Yes" Then
        frmCountries.Hide
        frmItalianVeggie.Show
    ElseIf veggie = "No" Then
        frmCountries.Hide
        frmItalian.Show
    Else: MsgBox ("Sorry, Please enter Yes or No.")
End If

End Sub

Private Sub cmdJapanese_Click(Index As Integer)
veggie = InputBox("Type in Yes if you are, No if you are not.", "Are you a vegetarian?")

If veggie = "Yes" Then
        frmCountries.Hide
        frmJapaneseVeggie.Show
    ElseIf veggie = "No" Then
        frmCountries.Hide
        frmJapanese.Show
    Else: MsgBox ("Sorry, Please enter Yes or No.")
End If

End Sub

Private Sub cmdMexican_Click(Index As Integer)
veggie = InputBox("Type in Yes if you are, No if you are not.", "Are you a vegetarian?")

If veggie = "Yes" Then
        frmCountries.Hide
        frmMexicanVeggie.Show
    ElseIf veggie = "No" Then
        frmCountries.Hide
        frmMexican.Show
    Else: MsgBox ("Sorry, Please enter Yes or No.")
End If

End Sub

Private Sub cmdQuit_Click(Index As Integer)
End
End Sub

