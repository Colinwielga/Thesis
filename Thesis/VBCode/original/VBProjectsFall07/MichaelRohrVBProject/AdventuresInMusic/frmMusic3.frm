VERSION 5.00
Begin VB.Form frmMusic3 
   Caption         =   "Music Basics 3"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   Picture         =   "frmMusic3.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh All Buttons"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5280
      Width           =   1335
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
      Height          =   1335
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      Picture         =   "frmMusic3.frx":243762
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   19
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox Picture10 
      Height          =   615
      Left            =   3360
      Picture         =   "frmMusic3.frx":244B94
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   1680
      Width           =   615
   End
   Begin VB.PictureBox Picture7 
      Height          =   735
      Left            =   4680
      Picture         =   "frmMusic3.frx":245E96
      ScaleHeight     =   675
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   5280
      Width           =   495
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2280
      Picture         =   "frmMusic3.frx":247198
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   16
      Top             =   5280
      Width           =   855
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5760
      Picture         =   "frmMusic3.frx":24817A
      ScaleHeight     =   795
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdHalf_Rest 
      BackColor       =   &H00FF80FF&
      Caption         =   "Half Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuarterRest1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Quarter Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdWholeRest1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Whole Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdWhole_Rest 
      BackColor       =   &H000080FF&
      Caption         =   "Whole Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuarterRest3 
      BackColor       =   &H000080FF&
      Caption         =   "Quarter Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdHalfRest1 
      BackColor       =   &H000080FF&
      Caption         =   "Half Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdEighthRest1 
      BackColor       =   &H0080FF80&
      Caption         =   "Eighth Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuarter_Rest 
      BackColor       =   &H0080FF80&
      Caption         =   "Quarter Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdSixteenthRest1 
      BackColor       =   &H0080FF80&
      Caption         =   "Sixteenth Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdEighthRest2 
      BackColor       =   &H00FF8080&
      Caption         =   "Eighth Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuarterRest2 
      BackColor       =   &H00FF8080&
      Caption         =   "Quarter Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdSixteenth_Rest 
      BackColor       =   &H00FF8080&
      Caption         =   "Sixteenth Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdEighth_Rest 
      BackColor       =   &H00FFFF00&
      Caption         =   "Eighth Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuarterRest4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Quarter Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdSixteenthRest2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Sixteenth Rest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   7680
      TabIndex        =   28
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Rests"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   3120
      TabIndex        =   25
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "6."
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
      Index           =   6
      Left            =   1920
      TabIndex        =   24
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbl9 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "7."
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
      Index           =   7
      Left            =   4080
      TabIndex        =   23
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbl10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "8."
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
      Index           =   8
      Left            =   6360
      TabIndex        =   22
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbl11 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "9."
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
      Index           =   9
      Left            =   3240
      TabIndex        =   21
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label lbl12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "10."
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
      Index           =   10
      Left            =   5280
      TabIndex        =   20
      Top             =   5280
      Width           =   495
   End
End
Attribute VB_Name = "frmMusic3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is a continuation of the previous form, frmMusic2, and tests the user on their knowledge of rests and asks 5 questions (the second set of 5 out of 15)

Private Sub cmdBack_Click()     'This button changes forms and goes back to the Main Page also setting the value of PointsMusic = 0
    frmMusic3.Hide                  'this hides frmMusic3
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
    PointsMusic = 0                 'this sets the value of PointsMusic = 0
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdEighth_Rest_Click()
    cmdQuarterRest4.Visible = False
    cmdSixteenthRest2.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdEighthRest1_Click()
    cmdQuarter_Rest.Visible = False
    cmdSixteenthRest1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdEighthRest2_Click()
    cmdQuarterRest2.Visible = False
    cmdSixteenth_Rest.Visible = False
End Sub

'This complicated button does so many things, like the one on frmMusic2 it uses If statements to tell the user via message box what questions are not yet answered
'it also uses Else statements if the questions are finished to increment a variable Num by 1 which is tested to know whether or not to change forms
'there are also If statements to give points checking if the correct button has been pushed and is remaining visible then adding a point to PointsMusic variable
Private Sub cmdFinished_Click()
Dim Num As Integer
    Num = 0                                                                                                 'sets value of Num = 0
    If cmdHalf_Rest.Visible = True And cmdQuarterRest1.Visible = True And cmdWholeRest1.Visible = True Then 'This set of If statements , like in frmMusic2, checks if the question has been answered and if not
        MsgBox "You Must Finish Question 6!", , "Unfinished Question"                                       'gives a message box telling the user to finish the question if so Num is incremented by 1,
    Else                                                                                                    'this is true of this entire set of If statements
        Num = Num + 1
    End If
    If cmdEighthRest1.Visible = True And cmdQuarter_Rest.Visible = True And cmdSixteenthRest1.Visible = True Then
        MsgBox "You Must Finish Question 7!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    If cmdEighthRest2.Visible = True And cmdQuarterRest2.Visible = True And cmdSixteenth_Rest.Visible = True Then
        MsgBox "You Must Finish Question 8!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    If cmdHalfRest1.Visible = True And cmdQuarterRest3.Visible = True And cmdWhole_Rest.Visible = True Then
        MsgBox "You Must Finish Question 9!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    If cmdEighth_Rest.Visible = True And cmdQuarterRest4.Visible = True And cmdSixteenthRest2.Visible = True Then
        MsgBox "You Must Finish Question 10!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
'This set of If statements, like in frmMusic2, checks if the command button clicked for the question is the correct answer,
'if so adds 1 point to the value of PointsMusic, if not nothing happens
    If cmdHalf_Rest.Visible = True And cmdQuarterRest1.Visible = False And cmdWholeRest1.Visible = False Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdEighthRest1.Visible = False And cmdQuarter_Rest.Visible = True And cmdSixteenthRest1.Visible = False Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdEighthRest2.Visible = False And cmdQuarterRest2.Visible = False And cmdSixteenth_Rest.Visible = True Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdHalfRest1.Visible = False And cmdQuarterRest3.Visible = False And cmdWhole_Rest.Visible = True Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdEighth_Rest.Visible = True And cmdQuarterRest4.Visible = False And cmdSixteenthRest2.Visible = False Then
        PointsMusic = PointsMusic + 1
    End If

'Here is where the Num variable comes back into play, if Num = 5 then user can move onto the next form,
'if not then a message box keeps them on this form
    If Num = 5 Then
        frmMusic3.Hide                  'this hides frmMusic3
        frmMusic4.Show                  'this makes frmMusic4 visible
        cmdHalf_Rest.Visible = True     'this places all of the buttons on this page to visible,
        cmdQuarterRest1.Visible = True  'otherwise if you came back to this page the buttons would already be clicked on
        cmdWholeRest1.Visible = True
        cmdEighthRest1.Visible = True
        cmdQuarter_Rest.Visible = True
        cmdSixteenthRest1.Visible = True
        cmdEighthRest2.Visible = True
        cmdQuarterRest2.Visible = True
        cmdSixteenth_Rest.Visible = True
        cmdHalfRest1.Visible = True
        cmdQuarterRest3.Visible = True
        cmdWhole_Rest.Visible = True
        cmdEighth_Rest.Visible = True
        cmdQuarterRest4.Visible = True
        cmdSixteenthRest2.Visible = True
    End If
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdHalf_Rest_Click()
    cmdQuarterRest1.Visible = False
    cmdWholeRest1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdHalfRest1_Click()
    cmdQuarterRest3.Visible = False
    cmdWhole_Rest.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdSixteenth_Rest_Click()
    cmdEighthRest2.Visible = False
    cmdQuarterRest2.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarter_Rest_Click()
    cmdEighthRest1.Visible = False
    cmdSixteenthRest1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarterRest1_Click()
    cmdHalf_Rest.Visible = False
    cmdWholeRest1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarterRest2_Click()
    cmdEighthRest2.Visible = False
    cmdSixteenth_Rest.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarterRest3_Click()
    cmdHalfRest1.Visible = False
    cmdQuarterRest3.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdQuarterRest4_Click()
    cmdEighth_Rest.Visible = False
    cmdSixteenthRest2.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdWholeRest1_Click()
    cmdHalf_Rest.Visible = False
    cmdQuarterRest1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdWhole_Rest_Click()
    cmdHalfRest1.Visible = False
    cmdQuarterRest3.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdSixteenthRest1_Click()
    cmdEighthRest1.Visible = False
    cmdQuarter_Rest.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdSixteenthRest2_Click()
    cmdEighth_Rest.Visible = False
    cmdQuarterRest4.Visible = False
End Sub

'This button makes all of the buttons on the form visible
'if a person made a mistake and wanted to redo a question this makes all of the buttons visible and they must be reclicked
'it does not interfer with the person's score whatsoever
Private Sub cmdRefresh_Click()
    cmdHalf_Rest.Visible = True
    cmdQuarterRest1.Visible = True
    cmdWholeRest1.Visible = True
    cmdEighthRest1.Visible = True
    cmdQuarter_Rest.Visible = True
    cmdSixteenthRest1.Visible = True
    cmdEighthRest2.Visible = True
    cmdQuarterRest2.Visible = True
    cmdSixteenth_Rest.Visible = True
    cmdHalfRest1.Visible = True
    cmdQuarterRest3.Visible = True
    cmdWhole_Rest.Visible = True
    cmdEighth_Rest.Visible = True
    cmdQuarterRest4.Visible = True
    cmdSixteenthRest2.Visible = True
End Sub
