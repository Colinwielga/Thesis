VERSION 5.00
Begin VB.Form frmMusic1 
   AutoRedraw      =   -1  'True
   Caption         =   "Music Basics Question 1"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   Picture         =   "frmMusic1.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdNatural 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   8640
      TabIndex        =   32
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdFlat 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   8640
      TabIndex        =   31
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSharp 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   8640
      TabIndex        =   30
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox picNeutral 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   9960
      Picture         =   "frmMusic1.frx":243762
      ScaleHeight     =   915
      ScaleWidth      =   555
      TabIndex        =   29
      Top             =   4560
      Width           =   615
   End
   Begin VB.PictureBox picFlat 
      Height          =   975
      Left            =   9960
      Picture         =   "frmMusic1.frx":24537C
      ScaleHeight     =   915
      ScaleWidth      =   555
      TabIndex        =   28
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox picSharp 
      Height          =   1095
      Left            =   9960
      Picture         =   "frmMusic1.frx":24762E
      ScaleHeight     =   1035
      ScaleWidth      =   675
      TabIndex        =   27
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdBassClef 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   7200
      TabIndex        =   26
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdTrebleClef 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   7200
      TabIndex        =   25
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picBass1 
      Height          =   1455
      Left            =   5280
      Picture         =   "frmMusic1.frx":24A688
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   24
      Top             =   4320
      Width           =   1695
   End
   Begin VB.PictureBox picTreble1 
      Height          =   2295
      Left            =   5280
      Picture         =   "frmMusic1.frx":25222A
      ScaleHeight     =   2235
      ScaleWidth      =   1515
      TabIndex        =   23
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdSixteenthR 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   2760
      TabIndex        =   22
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdEighthR 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   2760
      TabIndex        =   21
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuarterR 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2760
      TabIndex        =   20
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdHalfR 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2760
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdWholeR 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2760
      TabIndex        =   18
      Top             =   960
      Width           =   1095
   End
   Begin VB.PictureBox picSixteenthR 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4080
      Picture         =   "frmMusic1.frx":25D91C
      ScaleHeight     =   795
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   5400
      Width           =   495
   End
   Begin VB.PictureBox picQuarterR 
      Height          =   615
      Left            =   4080
      Picture         =   "frmMusic1.frx":25EA8E
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   3000
      Width           =   615
   End
   Begin VB.PictureBox picHalfR 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      Picture         =   "frmMusic1.frx":25FD90
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   15
      Top             =   1800
      Width           =   855
   End
   Begin VB.PictureBox picWholeR 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      Picture         =   "frmMusic1.frx":2611C2
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox picEighthR 
      Height          =   735
      Left            =   4080
      Picture         =   "frmMusic1.frx":2621A4
      ScaleHeight     =   675
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdSixteenth 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1320
      TabIndex        =   12
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdEighth 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   1320
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuarter 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdhalf 
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1320
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdWhole 
      BackColor       =   &H8000000D&
      Caption         =   "What Am I?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1320
      MaskColor       =   &H8000000D&
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.PictureBox picSixteenthN 
      Height          =   1095
      Left            =   480
      Picture         =   "frmMusic1.frx":2634A6
      ScaleHeight     =   1035
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   5520
      Width           =   615
   End
   Begin VB.PictureBox picQuarterN 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      Picture         =   "frmMusic1.frx":2656F0
      ScaleHeight     =   915
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox picEighthN 
      Height          =   1095
      Left            =   360
      Picture         =   "frmMusic1.frx":266B36
      ScaleHeight     =   1035
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   4200
      Width           =   735
   End
   Begin VB.PictureBox picWholeN 
      Height          =   615
      Left            =   240
      Picture         =   "frmMusic1.frx":2692C8
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdFinished 
      BackColor       =   &H0000FFFF&
      Caption         =   "Test Your Stuff"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   1575
   End
   Begin VB.PictureBox picHalfN 
      AutoSize        =   -1  'True
      Height          =   1035
      Left            =   480
      Picture         =   "frmMusic1.frx":26B146
      ScaleHeight     =   975
      ScaleWidth      =   510
      TabIndex        =   2
      Top             =   1800
      Width           =   570
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back to Main Page"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label lblColor4 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   8520
      TabIndex        =   36
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblColor3 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   5160
      TabIndex        =   35
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblColor2 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   2640
      TabIndex        =   34
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblColor1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   33
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblQuestionM1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Music Matching"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmMusic1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The purpose of this form is to inform the user of Musical concepts using message boxes
'that spring up by clicking the many many buttons on the page and then leading to a quiz that can be taken to asses their knowledge

'This button displays a message box for the bass clef
Private Sub cmdBassClef_Click(Index As Integer)
    MsgBox "I am a Bass Clef, my lines spell 'G, B, D, F, A' and my spaces spell 'A, C, E, G'.", , "What Am I?"
End Sub

Private Sub cmdContinue_Click()     'This button changes forms back to the main page
    frmMusic1.Hide                  'this hides frmMusic1
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
End Sub

'This button displays a message box for the eighth note
Private Sub cmdEighth_Click(Index As Integer)
    MsgBox "I am an Eighth Note, I am one half of a beat long.", , "What Am I?"
End Sub

'This button displays a message box for the eighth rest
Private Sub cmdEighthR_Click(Index As Integer)
    MsgBox "I am a Eighth Rest, I am silent for one half a beat.", , "What Am I?"
End Sub

Private Sub cmdFinished_Click()     'This button changes forms leading to the quiz
    frmMusic1.Hide                      'this hides frmMusic1
    frmMusic2.Show                      'this makes frmMusic2 visible
End Sub

'This button displays a message box for the flat sign
Private Sub cmdFlat_Click(Index As Integer)
    MsgBox "I am a Flat, I lower a pitch by a half step.", , "What Am I?"
End Sub

'This button displays a message box for the half note
Private Sub cmdhalf_Click(Index As Integer)
    MsgBox "I am a Half Note, I am two beats long.", , "What Am I?"
End Sub

'This button displays a message box for the half rest
Private Sub cmdHalfR_Click(Index As Integer)
    MsgBox "I am a Half Rest, I am silent for two beats.", , "What Am I?"
End Sub

'This button displays a message box for the natural sign
Private Sub cmdNatural_Click(Index As Integer)
    MsgBox "I am a Natural, I return any sharp or flat back to its natural pitch.", , "What Am I?"
End Sub

'This button displays a message box for the quarter note
Private Sub cmdQuarter_Click(Index As Integer)
    MsgBox "I am a Quarter Note, I am one beat long.", , "What Am I?"
End Sub

'This button displays a message box for the quarter rest
Private Sub cmdQuarterR_Click(Index As Integer)
    MsgBox "I am a Quarter Rest, I am silent for one beat.", , "What Am I?"
End Sub

'This button displays a message box for the sharp sign
Private Sub cmdSharp_Click(Index As Integer)
    MsgBox "I am a Sharp, I raise a pitch by a half step.", , "What Am I?"
End Sub

'This button displays a message box for the sixteenth note
Private Sub cmdSixteenth_Click(Index As Integer)
    MsgBox "I am a Sixteenth Note, I am one fourth of a beat long.", , "What Am I?"
End Sub

'This button displays a message box for the sixteenth rest
Private Sub cmdSixteenthR_Click(Index As Integer)
    MsgBox "I am a Sixteenth Rest, I am silent for one fourth of a beat.", , "What Am I?"
End Sub

'This button displays a message box for the treble clef
Private Sub cmdTrebleClef_Click(Index As Integer)
    MsgBox "I am a Treble Clef, my lines spell 'E, G, B, D, F' and my spaces spell 'F, A, C, E'.", , "What Am I?"
End Sub

'This button displays a message box for the whole note
Private Sub cmdWhole_Click(Index As Integer)
    MsgBox "I am a Whole Note, I am four beats long.", , "What Am I?"
End Sub

'This button displays a message box for the whole rest
Private Sub cmdWholeR_Click(Index As Integer)
    MsgBox " I am a Whole Rest, I am silent for four beats.", , "What Am I?"
End Sub
