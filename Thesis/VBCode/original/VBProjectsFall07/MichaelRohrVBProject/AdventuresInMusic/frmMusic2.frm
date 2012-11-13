VERSION 5.00
Begin VB.Form frmMusic2 
   BackColor       =   &H00000000&
   Caption         =   "Music Basics 2"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmMusic2.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuarterNote4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quarter Note"
      Height          =   615
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdEighthNote2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Eighth Note"
      Height          =   615
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdSixteenth_Note 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sixteenth Note"
      Height          =   615
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdEighthNote1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Eighth Note"
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdHalfNote2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Half Note"
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuarter_Note 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quarter Note"
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuarterNote2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quarter Note"
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdEighth_Note 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Eighth Note"
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdSixteenthNote1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sixteenth Note"
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdHalfNote1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Half Note"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuarterNote3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Quarter Note"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdWhole_Note 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Whole Note"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh All Buttons"
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdHalf_Note 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Half Note"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuarterNote1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quarter Note"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdWholeNote1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Whole Note"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   1815
   End
   Begin VB.PictureBox Picture5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9480
      Picture         =   "frmMusic2.frx":243762
      ScaleHeight     =   1035
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   2880
      Width           =   615
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      Picture         =   "frmMusic2.frx":2459AC
      ScaleHeight     =   915
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Picture         =   "frmMusic2.frx":246DF2
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      Picture         =   "frmMusic2.frx":248C70
      ScaleHeight     =   1035
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      Picture         =   "frmMusic2.frx":24B402
      ScaleHeight     =   975
      ScaleWidth      =   510
      TabIndex        =   4
      Top             =   3000
      Width           =   570
   End
   Begin VB.CommandButton cmdFinished 
      BackColor       =   &H0080FFFF&
      Caption         =   "Finished"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
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
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   30
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   10200
      TabIndex        =   13
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   12
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   11
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   7
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblQuestionM2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Music Basics Lesson"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblMusicQuestion2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   $"frmMusic2.frx":24CEAC
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1695
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   7575
   End
End
Attribute VB_Name = "frmMusic2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This page is the first of three form pages that quiz the user on the concepts of music
'this page focuses on notes and uses command buttons to answer the questions by clicking on one and making invisible the others asking 5 questions (15 total in the quiz),
'there is a refresh button that makes all the buttons on the form visible
'the finished button uses if statesments to tell if the question is both answered and if the answer is correct,
'if so a point is incremented to the Public variable PointsMusic
'Also I got the idea for making command buttons visible, not visible to create solutions from the Hogwarts example project,
'I didn't know if I needed to reference that or not.

Private Sub cmdBack_Click()     'This button changes forms to frmLessonMainPage and clears and sets value of PointsMusic = 0
    frmMusic2.Hide                  'this hides frmMusic2
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
    PointsMusic = 0                 'this sets PointsMusic = 0
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdEighth_Note_Click()
    cmdSixteenthNote1.Visible = False
    cmdQuarterNote2.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdEighthNote1_Click()
    cmdQuarter_Note.Visible = False
    cmdHalfNote2.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdEighthNote2_Click()
    cmdSixteenth_Note.Visible = False
    cmdQuarterNote4.Visible = False
End Sub

'This complicated button does so many things, it uses a plethera of If statements to tell the user via message box what questions might not be answered
'it also uses Else statements if the questions are finished to increment a variable Num by 1 which is tested to know whether or not to change forms
'there are also If statements to give points checking if the correct button has been pushed and is remaining visible then adding a point to PointsMusic variable
Private Sub cmdFinished_Click()
Dim Num As Integer
    Num = 0                                                                                                 'set Num = 0
    If cmdWholeNote1.Visible = True And cmdQuarterNote1.Visible = True And cmdHalf_Note.Visible = True Then 'This set of If statements checks if the question has been answered and if not gives a message box telling the user to finish the question
        MsgBox "You Must Finish Question 1!", , "Unfinished Question"                                       'if so Num is incremented by 1, this is true of this entire set of If statements
    Else
        Num = Num + 1
    End If
    If cmdSixteenthNote1.Visible = True And cmdEighth_Note.Visible = True And cmdQuarterNote2.Visible = True Then
        MsgBox "You Must Finish Question 2!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    If cmdHalfNote1.Visible = True And cmdQuarterNote3.Visible = True And cmdWhole_Note.Visible = True Then
        MsgBox "You Must Finish Question 3!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    If cmdQuarter_Note.Visible = True And cmdHalfNote2.Visible = True And cmdEighthNote1.Visible = True Then
        MsgBox "You Must Finish Question 4!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    If cmdSixteenth_Note.Visible = True And cmdEighthNote2.Visible = True And cmdQuarterNote4.Visible = True Then
        MsgBox "You Must Finish Question 5!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    PointsMusic = 0                                                                                             'sets value of PointsMusic = 0
    If cmdWholeNote1.Visible = False And cmdQuarterNote1.Visible = False And cmdHalf_Note.Visible = True Then   'This set of If statement checks to see if the correct button is visible, and if so adds a point to PointsMusic, if not it does nothing
        PointsMusic = PointsMusic + 1                                                                           'This is true of all of these If statements here
    End If
    If cmdSixteenthNote1.Visible = False And cmdEighth_Note.Visible = True And cmdQuarterNote2.Visible = False Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdHalfNote1.Visible = False And cmdQuarterNote3.Visible = False And cmdWhole_Note.Visible = True Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdQuarter_Note.Visible = True And cmdHalfNote2.Visible = False And cmdEighthNote1.Visible = False Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdSixteenth_Note.Visible = True And cmdEighthNote2.Visible = False And cmdQuarterNote4.Visible = False Then
        PointsMusic = PointsMusic + 1
    End If
    If Num = 5 Then                     'Here is where the Num variable comes back into play, if Num = 5 then user can move onto the next form, if not then a message box keeps them on this form, before I put this in the pages would still switch this was the only way I could think of to keep that from happening
        frmMusic2.Hide                  'this hides frmMusic2
        frmMusic3.Show                  'this makes frmMusic3 visible
        cmdHalf_Note.Visible = True     'this places all of the buttons on this page to visible,
        cmdQuarterNote1.Visible = True  'otherwise if you came back to this page the buttons would already be clicked on
        cmdWholeNote1.Visible = True
        cmdSixteenthNote1.Visible = True
        cmdEighth_Note.Visible = True
        cmdQuarterNote2.Visible = True
        cmdWhole_Note.Visible = True
        cmdQuarterNote3.Visible = True
        cmdHalfNote1.Visible = True
        cmdQuarter_Note.Visible = True
        cmdHalfNote2.Visible = True
        cmdEighthNote1.Visible = True
        cmdSixteenth_Note.Visible = True
        cmdEighthNote2.Visible = True
        cmdQuarterNote4.Visible = True
    End If
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdHalf_Note_Click()
    cmdWholeNote1.Visible = False
    cmdQuarterNote1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdHalfNote1_Click()
    cmdWhole_Note.Visible = False
    cmdQuarterNote3.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdHalfNote2_Click()
    cmdQuarter_Note.Visible = False
    cmdEighthNote1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarter_Note_Click()
    cmdHalfNote2.Visible = False
    cmdEighthNote1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarterNote1_Click()
    cmdHalf_Note.Visible = False
    cmdWholeNote1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarterNote2_Click()
    cmdSixteenthNote1.Visible = False
    cmdEighth_Note.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarterNote3_Click()
    cmdWhole_Note.Visible = False
    cmdHalfNote1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarterNote4_Click()
    cmdSixteenth_Note.Visible = False
    cmdEighthNote2.Visible = False
End Sub

'This button makes all of the buttons on the form visible
'if a person made a mistake and wanted to redo a question this makes all of the buttons visible and they must be reclicked
Private Sub cmdRefresh_Click()
    cmdHalf_Note.Visible = True
    cmdQuarterNote1.Visible = True
    cmdWholeNote1.Visible = True
    cmdSixteenthNote1.Visible = True
    cmdEighth_Note.Visible = True
    cmdQuarterNote2.Visible = True
    cmdWhole_Note.Visible = True
    cmdQuarterNote3.Visible = True
    cmdHalfNote1.Visible = True
    cmdQuarter_Note.Visible = True
    cmdHalfNote2.Visible = True
    cmdEighthNote1.Visible = True
    cmdSixteenth_Note.Visible = True
    cmdEighthNote2.Visible = True
    cmdQuarterNote4.Visible = True
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdSixteenth_Note_Click()
    cmdEighthNote2.Visible = False
    cmdQuarterNote4.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdSixteenthNote1_Click()
    cmdEighth_Note.Visible = False
    cmdQuarterNote2.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdWhole_Note_Click()
    cmdQuarterNote3.Visible = False
    cmdHalfNote1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdWholeNote1_Click()
    cmdQuarterNote1.Visible = False
    cmdHalf_Note.Visible = False
End Sub

