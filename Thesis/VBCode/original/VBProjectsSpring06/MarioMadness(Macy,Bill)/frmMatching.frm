VERSION 5.00
Begin VB.Form frmMatching 
   Caption         =   "Mario Matching Game"
   ClientHeight    =   10140
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   Picture         =   "frmMatching.frx":0000
   ScaleHeight     =   10140
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset the buttons"
      Height          =   852
      Left            =   6360
      TabIndex        =   74
      Top             =   360
      Width           =   1932
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   18
      Left            =   9120
      TabIndex        =   31
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   17
      Left            =   7440
      TabIndex        =   33
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   16
      Left            =   5760
      TabIndex        =   34
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   14
      Left            =   2400
      TabIndex        =   32
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   15
      Left            =   4080
      TabIndex        =   30
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   10
      Left            =   5760
      TabIndex        =   17
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   9
      Left            =   4080
      TabIndex        =   18
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   8
      Left            =   2400
      TabIndex        =   19
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   1
      Left            =   720
      TabIndex        =   13
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   2
      Left            =   2400
      TabIndex        =   26
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   3
      Left            =   4080
      TabIndex        =   25
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   6
      Left            =   9120
      TabIndex        =   15
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   5
      Left            =   7440
      TabIndex        =   23
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   17
      Left            =   7440
      TabIndex        =   28
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   16
      Left            =   5760
      TabIndex        =   29
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   15
      Left            =   4080
      TabIndex        =   21
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   14
      Left            =   2400
      TabIndex        =   12
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   13
      Left            =   720
      TabIndex        =   14
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   12
      Left            =   9120
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   11
      Left            =   7440
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   10
      Left            =   5760
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   9
      Left            =   4080
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   8
      Left            =   2400
      TabIndex        =   11
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   12
      Left            =   9120
      TabIndex        =   35
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   7
      Left            =   720
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   6
      Left            =   9120
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   5
      Left            =   7440
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   11
      Left            =   7440
      TabIndex        =   16
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   4
      Left            =   5760
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   13
      Left            =   720
      TabIndex        =   22
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   3
      Left            =   4080
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   7
      Left            =   720
      TabIndex        =   20
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   4
      Left            =   5760
      TabIndex        =   24
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Click me!"
      Height          =   1215
      Index           =   18
      Left            =   9120
      TabIndex        =   27
      Top             =   4440
      Width           =   1455
   End
   Begin VB.PictureBox pic17 
      Height          =   852
      Left            =   7800
      Picture         =   "frmMatching.frx":1DDB1
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   71
      Top             =   1800
      Width           =   852
   End
   Begin VB.PictureBox pic21 
      Height          =   852
      Left            =   4320
      Picture         =   "frmMatching.frx":1E265
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   70
      Top             =   6120
      Width           =   852
   End
   Begin VB.PictureBox pic24 
      Height          =   852
      Left            =   9360
      Picture         =   "frmMatching.frx":1E719
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   69
      Top             =   6120
      Width           =   972
   End
   Begin VB.PictureBox pic13 
      Height          =   852
      Left            =   4320
      Picture         =   "frmMatching.frx":1EC5F
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   68
      Top             =   3240
      Width           =   972
   End
   Begin VB.PictureBox pic34 
      Height          =   852
      Left            =   4320
      Picture         =   "frmMatching.frx":1F1A5
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   67
      Top             =   4680
      Width           =   972
   End
   Begin VB.PictureBox pic26 
      Height          =   852
      Left            =   2640
      Picture         =   "frmMatching.frx":1F706
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   66
      Top             =   7560
      Width           =   972
   End
   Begin VB.PictureBox pic5 
      Height          =   852
      Left            =   7680
      Picture         =   "frmMatching.frx":1FC67
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   65
      Top             =   4680
      Width           =   972
   End
   Begin VB.PictureBox pic27 
      Height          =   852
      Left            =   4320
      Picture         =   "frmMatching.frx":2020D
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   64
      Top             =   7560
      Width           =   972
   End
   Begin VB.PictureBox pic30 
      Height          =   852
      Left            =   9480
      Picture         =   "frmMatching.frx":207B3
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   63
      Top             =   7560
      Width           =   852
   End
   Begin VB.PictureBox pic7 
      Height          =   852
      Left            =   960
      Picture         =   "frmMatching.frx":20CBB
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   62
      Top             =   3240
      Width           =   852
   End
   Begin VB.PictureBox pic35 
      Height          =   852
      Left            =   9480
      Picture         =   "frmMatching.frx":211C3
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   61
      Top             =   4680
      Width           =   852
   End
   Begin VB.PictureBox pic19 
      Height          =   852
      Left            =   960
      Picture         =   "frmMatching.frx":21745
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   60
      Top             =   6120
      Width           =   852
   End
   Begin VB.PictureBox pic12 
      Height          =   852
      Left            =   9360
      Picture         =   "frmMatching.frx":21CC7
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   59
      Top             =   3240
      Width           =   972
   End
   Begin VB.PictureBox pic15 
      Height          =   852
      Left            =   6000
      Picture         =   "frmMatching.frx":22258
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   58
      Top             =   9000
      Width           =   972
   End
   Begin VB.PictureBox pic10 
      Height          =   852
      Left            =   6000
      Picture         =   "frmMatching.frx":227E9
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   57
      Top             =   3240
      Width           =   852
   End
   Begin VB.PictureBox pic29 
      Height          =   852
      Left            =   7800
      Picture         =   "frmMatching.frx":22E71
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   56
      Top             =   7560
      Width           =   852
   End
   Begin VB.PictureBox pic32 
      Height          =   852
      Left            =   2640
      Picture         =   "frmMatching.frx":2353D
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   55
      Top             =   9000
      Width           =   852
   End
   Begin VB.PictureBox pic6 
      Height          =   852
      Left            =   9360
      Picture         =   "frmMatching.frx":23BC5
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   53
      Top             =   1800
      Width           =   972
   End
   Begin VB.PictureBox pic28 
      Height          =   852
      Left            =   6000
      Picture         =   "frmMatching.frx":271C5
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   52
      Top             =   7560
      Width           =   972
   End
   Begin VB.PictureBox pic20 
      Height          =   852
      Left            =   2640
      Picture         =   "frmMatching.frx":2A7C5
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   51
      Top             =   4680
      Width           =   972
   End
   Begin VB.PictureBox pic23 
      Height          =   852
      Left            =   7680
      Picture         =   "frmMatching.frx":2DB4C
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   50
      Top             =   6120
      Width           =   972
   End
   Begin VB.PictureBox pic8 
      Height          =   852
      Left            =   2640
      Picture         =   "frmMatching.frx":30ED3
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   49
      Top             =   3240
      Width           =   972
   End
   Begin VB.PictureBox pic18 
      Height          =   852
      Left            =   7680
      Picture         =   "frmMatching.frx":343DC
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   48
      Top             =   9000
      Width           =   972
   End
   Begin VB.PictureBox pic33 
      Height          =   852
      Left            =   4320
      Picture         =   "frmMatching.frx":378E5
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   47
      Top             =   9000
      Width           =   972
   End
   Begin VB.PictureBox pic16 
      Height          =   852
      Left            =   6000
      Picture         =   "frmMatching.frx":3BACD
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   46
      Top             =   4680
      Width           =   972
   End
   Begin VB.PictureBox pic14 
      Height          =   852
      Left            =   2640
      Picture         =   "frmMatching.frx":3FCB5
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   45
      Top             =   6120
      Width           =   972
   End
   Begin VB.PictureBox pic11 
      Height          =   852
      Left            =   7680
      Picture         =   "frmMatching.frx":43FDB
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   44
      Top             =   3240
      Width           =   972
   End
   Begin VB.PictureBox pic36 
      Height          =   852
      Left            =   9360
      Picture         =   "frmMatching.frx":48301
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   43
      Top             =   9000
      Width           =   972
   End
   Begin VB.PictureBox pic9 
      Height          =   852
      Left            =   960
      Picture         =   "frmMatching.frx":4BAD6
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   42
      Top             =   4680
      Width           =   972
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Return to the main page"
      Height          =   855
      Left            =   8640
      TabIndex        =   36
      Top             =   360
      Width           =   1932
   End
   Begin VB.PictureBox pic4 
      Height          =   852
      Left            =   6120
      Picture         =   "frmMatching.frx":4F2AB
      ScaleHeight     =   804
      ScaleWidth      =   804
      TabIndex        =   54
      Top             =   1800
      Width           =   852
   End
   Begin VB.PictureBox pic1 
      Height          =   852
      Left            =   960
      Picture         =   "frmMatching.frx":4F977
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   38
      Top             =   1800
      Width           =   972
   End
   Begin VB.PictureBox pic2 
      Height          =   852
      Left            =   2640
      Picture         =   "frmMatching.frx":51B82
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   40
      Top             =   1800
      Width           =   972
   End
   Begin VB.PictureBox pic3 
      Height          =   852
      Left            =   4320
      Picture         =   "frmMatching.frx":5289E
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   72
      Top             =   1800
      Width           =   972
   End
   Begin VB.PictureBox pic22 
      Height          =   852
      Left            =   6000
      Picture         =   "frmMatching.frx":580DA
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   39
      Top             =   6120
      Width           =   972
   End
   Begin VB.PictureBox pic25 
      Height          =   852
      Left            =   960
      Picture         =   "frmMatching.frx":5A2E5
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   41
      Top             =   7560
      Width           =   972
   End
   Begin VB.PictureBox pic31 
      Height          =   852
      Left            =   960
      Picture         =   "frmMatching.frx":5B001
      ScaleHeight     =   804
      ScaleWidth      =   924
      TabIndex        =   73
      Top             =   9000
      Width           =   972
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   75
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   240
      X2              =   11040
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label lblMatching 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mario Matching Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   23.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   37
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmMatching
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to play a great game of matching.  The use can select an image from the top
                'and then select one from the bottom to see if they match.  If the do, they stay uncovered and if they dont
                'they are covered back up after a message is displayed.  They also have the option to reset the game to start
                'over or to return to the main page.


Option Explicit
Dim lastcmdpic1Clicked As Integer       'declares all my variables
Dim lastcmdpic2Clicked As Integer
Dim Mapping(1 To 18) As Integer
Dim index As Integer

Private Sub cmdexit_Click()
    frmMain.Show        'shows the main form
    frmMatching.Hide        'hides the matching form
End Sub


Private Sub cmdpic1_Click(index As Integer)
    lastcmdpic1Clicked = index      'when a picture from the top is selected the index is stored in this variable
    cmdpic1(index).Visible = False      'the picture becomes visible to the user
End Sub

Private Sub cmdpic2_Click(index As Integer)
    
    lastcmdpic2Clicked = index      'the index of the second picture is stored in this variable
    cmdpic2(index).Visible = False  'the second half picture is shown to the user
    If cmdpic1(Mapping(index)).Visible = True Then      'if the mapped index of the first pic is the same as the second stored below it loops
        cmdpic2(index).Visible = True       'picture two becomes visible (button is hid)
        cmdpic1(lastcmdpic1Clicked).Visible = True      'picture one remains visible (button is hid)
    End If

    If cmdpic2(index).Visible = True Then       'if the second picture is hidden because it doesnt match this loop is followed
        cmdpic2(index).Visible = False      'the user gets a look at the second picture
        MsgBox "Nice try", , "Try again"        'a message is displayed telling them it was not a match
        cmdpic2(index).Visible = True       'the second picture is hidden from the user
    End If
End Sub

Private Sub cmdReset_Click()
    cmdpic2(1).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(1).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(2).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(2).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(3).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(3).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(4).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(4).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(5).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(5).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(6).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(6).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(7).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(7).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(8).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(8).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(9).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(9).Visible = True       'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(10).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(10).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(11).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(11).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(12).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(12).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(13).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(13).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(14).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(14).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(15).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(15).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(16).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(16).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(17).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(17).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic2(18).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    cmdpic1(18).Visible = True      'hides the button so that the user can play again.  This sets the visibility of the button to true so the button hides the picture again
    
End Sub

Private Sub Form_Load()
    
    Mapping(4) = 1      'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(7) = 2      'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(13) = 3     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(11) = 4     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(3) = 5      'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(10) = 6     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(12) = 7     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(17) = 8     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(6) = 9      'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(14) = 10        'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(2) = 11     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(16) = 12        'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(18) = 13        'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(5) = 14     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(8) = 15     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(15) = 16        'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(9) = 17     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
    Mapping(1) = 18     'sets the top picture and bottom picture equal so if they are clicked it will stay uncovered
End Sub

