VERSION 5.00
Begin VB.Form frmMusic4 
   Caption         =   "Music Basics 4"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   Picture         =   "frmMusic4.frx":0000
   ScaleHeight     =   9015
   ScaleWidth      =   9750
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdTreble1 
      BackColor       =   &H00FF00FF&
      Caption         =   "Treble Clef"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdBass 
      BackColor       =   &H00FF00FF&
      Caption         =   "Bass Clef"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdTreble 
      BackColor       =   &H000000FF&
      Caption         =   "Treble Clef"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdBass1 
      BackColor       =   &H000000FF&
      Caption         =   "Bass Clef"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7440
      Width           =   1815
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture12 
      Height          =   1455
      Left            =   4320
      Picture         =   "frmMusic4.frx":243762
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   16
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox Picture11 
      Height          =   2295
      Left            =   480
      Picture         =   "frmMusic4.frx":24B304
      ScaleHeight     =   2235
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   5400
      Width           =   1575
   End
   Begin VB.PictureBox Picture13 
      Height          =   975
      Left            =   5520
      Picture         =   "frmMusic4.frx":2569F6
      ScaleHeight     =   915
      ScaleWidth      =   675
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1200
      Picture         =   "frmMusic4.frx":259A50
      ScaleHeight     =   915
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox Picture14 
      Height          =   975
      Left            =   3360
      Picture         =   "frmMusic4.frx":25B66A
      ScaleHeight     =   915
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdSharp1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sharp"
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdFlat1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Flat"
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdNatural_1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Natural"
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdSharp2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Sharp"
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
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdFlat_2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Flat"
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
      TabIndex        =   4
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdNatural2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Natural"
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
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdSharp_3 
      BackColor       =   &H008080FF&
      Caption         =   "Sharp"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdFlat3 
      BackColor       =   &H008080FF&
      Caption         =   "Flat"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdNatural3 
      BackColor       =   &H008080FF&
      Caption         =   "Natural"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   7920
      TabIndex        =   26
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "15."
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
      Left            =   6240
      TabIndex        =   24
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "14."
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
      Left            =   2160
      TabIndex        =   23
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Signs and Clefs"
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
      Left            =   2760
      TabIndex        =   17
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "11."
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
      Index           =   4
      Left            =   1920
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "12."
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
      Index           =   5
      Left            =   4080
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl13 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "13."
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
      Index           =   11
      Left            =   6360
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "frmMusic4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form page is a continuation of frmMusic3, which is a continuation of frmMusic2, and quizes the user's knowledge of signs and clefs and asks 5 questions (the third set of 5 out of 15)

'This button makes one command button invisible leaving this one that was clicked visible
Private Sub cmdBass_Click()
    cmdTreble1.Visible = False
End Sub

'This button makes one command button invisible leaving this one that was clicked visible
Private Sub cmdBass1_Click()
    cmdTreble.Visible = False
End Sub

'This complicated button does so many things, like the one on frmMusic2 and frmMusic3, it uses If statements to tell the user via message box what questions are not yet answered
'it also uses Else statements if the questions are finished to increment a variable Num by 1 which is tested to know whether or not to change forms
'there are also If statements to give points checking if the correct button has been pushed and is remaining visible then adding a point to PointsMusic variable
'the very last things it does is make all of the buttons visible and then takes the user back to frmLessonMainPage
Private Sub cmdFinished_Click()
Dim Num As Integer
    Num = 0                                                                                         'set Num = 0
    If cmdSharp1.Visible = True And cmdFlat1.Visible = True And cmdNatural_1.Visible = True Then    'This set of If statements , like in frmMusic2 and frmMusic3, checks if the question has been answered and if not gives a message box telling them to finish the question
        MsgBox "You Must Finish Question 11!", , "Unfinished Question"                              'if so Num is incremented by 1, this is true of this entire set of If statements
    Else                                                                                            'gives a message box telling the user to finish the question if so Num is incremented by 1,
        Num = Num + 1                                                                               'this is true of this entire set of If statements
    End If
    If cmdSharp2.Visible = True And cmdFlat_2.Visible = True And cmdNatural2.Visible = True Then
        MsgBox "You Must Finish Question 12!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    If cmdSharp_3.Visible = True And cmdFlat3.Visible = True And cmdNatural3.Visible = True Then
        MsgBox "You Must Finish Question 13!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    If cmdTreble.Visible = True And cmdBass1.Visible = True Then
        MsgBox "You Must Finish Question 14!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
    If cmdTreble1.Visible = True And cmdBass.Visible = True Then
        MsgBox "You Must Finish Question 15!", , "Unfinished Question"
    Else
        Num = Num + 1
    End If
'This set of If statement checks to see if the correct button is visible, and if so adds a point to PointsMusic, if not it does nothing
'This is true of all of these If statements here
    If cmdSharp1.Visible = False And cmdFlat1.Visible = False And cmdNatural_1.Visible = True Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdSharp2.Visible = False And cmdFlat_2.Visible = True And cmdNatural2.Visible = False Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdSharp_3.Visible = True And cmdFlat3.Visible = False And cmdNatural3.Visible = False Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdTreble.Visible = True And cmdBass1.Visible = False Then
        PointsMusic = PointsMusic + 1
    End If
    If cmdTreble1.Visible = False And cmdBass.Visible = True Then
        PointsMusic = PointsMusic + 1
    End If
'Here is where the Num variable comes back into play, if Num = 5 then user can move onto the next form,
'if not then a message box keeps them on this form, when they have finished a message box pops up and displays there score and there name given on frmOpeningPage
    If Num = 5 Then
        MsgBox "Congratulations " & NameGiven & " You got " & PointsMusic & " points!!!  Way to go!", , "Your Score"
        frmMusic4.Hide              'this hides frmMusic4
        frmLessonMainPage.Show      'this makes frmLessonMainPage visible
        cmdSharp1.Visible = True    'this makes all of the buttons visible so that when the user comes back the buttons are no longer clicked
        cmdFlat1.Visible = True
        cmdNatural_1.Visible = True
        cmdFlat_2.Visible = True
        cmdSharp2.Visible = True
        cmdNatural2.Visible = True
        cmdSharp_3.Visible = True
        cmdFlat3.Visible = True
        cmdNatural3.Visible = True
        cmdTreble.Visible = True
        cmdTreble1.Visible = True
        cmdBass.Visible = True
        cmdBass1.Visible = True
    End If
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdFlat_2_Click()
    cmdSharp2.Visible = False
    cmdNatural2.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdFlat1_Click()
    cmdSharp1.Visible = False
    cmdNatural_1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdFlat3_Click()
    cmdSharp_3.Visible = False
    cmdNatural3.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdNatural_1_Click()
    cmdSharp1.Visible = False
    cmdFlat1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdNatural2_Click()
    cmdSharp2.Visible = False
    cmdFlat_2.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdNatural3_Click()
    cmdSharp_3.Visible = False
    cmdFlat3.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdSharp_3_Click()
    cmdFlat3.Visible = False
    cmdNatural3.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdSharp1_Click()
    cmdFlat1.Visible = False
    cmdNatural_1.Visible = False
End Sub

'This button makes two command buttons invisible leaving this one that was clicked visible
Private Sub cmdSharp2_Click()
    cmdFlat_2.Visible = False
    cmdNatural2.Visible = False
End Sub

'This button makes all of the buttons on the form visible
'if a person made a mistake and wanted to redo a question this makes all of the buttons visible and they must be reclicked
'it does not interfer with the person's score whatsoever
Private Sub cmdRefresh_Click()
    cmdSharp1.Visible = True
    cmdFlat1.Visible = True
    cmdNatural_1.Visible = True
    cmdFlat_2.Visible = True
    cmdSharp2.Visible = True
    cmdNatural2.Visible = True
    cmdSharp_3.Visible = True
    cmdFlat3.Visible = True
    cmdNatural3.Visible = True
    cmdTreble.Visible = True
    cmdTreble1.Visible = True
    cmdBass.Visible = True
    cmdBass1.Visible = True
End Sub

'This button makes one command button invisible leaving this one that was clicked visible
Private Sub cmdTreble_Click()
    cmdBass1.Visible = False
End Sub

'This button makes one command button invisible leaving this one that was clicked visible
Private Sub cmdTreble1_Click()
    cmdBass.Visible = False
End Sub

