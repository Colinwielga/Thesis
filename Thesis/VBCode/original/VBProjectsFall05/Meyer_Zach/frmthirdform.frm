VERSION 5.00
Begin VB.Form frmthirdform 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Take 'Em"
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   Picture         =   "frmthirdform.frx":0000
   ScaleHeight     =   10905
   ScaleWidth      =   14460
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picName 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      ScaleHeight     =   375
      ScaleWidth      =   2415
      TabIndex        =   16
      Top             =   9240
      Width           =   2415
   End
   Begin VB.PictureBox pic12 
      Height          =   3375
      Left            =   9240
      Picture         =   "frmthirdform.frx":1D1BC
      ScaleHeight     =   3315
      ScaleWidth      =   4515
      TabIndex        =   14
      Top             =   5160
      Width           =   4575
   End
   Begin VB.PictureBox pic9 
      Height          =   1575
      Left            =   4920
      Picture         =   "frmthirdform.frx":21549
      ScaleHeight     =   1515
      ScaleWidth      =   3195
      TabIndex        =   13
      Top             =   3240
      Width           =   3255
   End
   Begin VB.PictureBox pic11 
      Height          =   1455
      Left            =   6120
      Picture         =   "frmthirdform.frx":2220D
      ScaleHeight     =   1395
      ScaleWidth      =   2235
      TabIndex        =   12
      Top             =   7200
      Width           =   2295
   End
   Begin VB.PictureBox pic10 
      Height          =   2295
      Left            =   10440
      Picture         =   "frmthirdform.frx":234F4
      ScaleHeight     =   2235
      ScaleWidth      =   3075
      TabIndex        =   11
      Top             =   360
      Width           =   3135
   End
   Begin VB.PictureBox pic5 
      Height          =   1455
      Left            =   8520
      Picture         =   "frmthirdform.frx":253BC
      ScaleHeight     =   1395
      ScaleWidth      =   2235
      TabIndex        =   10
      Top             =   3480
      Width           =   2295
   End
   Begin VB.PictureBox pic8 
      Height          =   1455
      Left            =   5640
      Picture         =   "frmthirdform.frx":264C0
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   5160
      Width           =   1935
   End
   Begin VB.PictureBox pic7 
      Height          =   1815
      Left            =   360
      Picture         =   "frmthirdform.frx":2726E
      ScaleHeight     =   1755
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   5160
      Width           =   1455
   End
   Begin VB.PictureBox pic3 
      Height          =   1455
      Left            =   5400
      Picture         =   "frmthirdform.frx":2781F
      ScaleHeight     =   1395
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox pic6 
      Height          =   1935
      Left            =   8400
      Picture         =   "frmthirdform.frx":27F32
      ScaleHeight     =   1875
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox pic4 
      Height          =   3015
      Left            =   2040
      Picture         =   "frmthirdform.frx":286B0
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   5040
      Width           =   3015
   End
   Begin VB.PictureBox pic2 
      Height          =   2535
      Left            =   1440
      Picture         =   "frmthirdform.frx":292B5
      ScaleHeight     =   2475
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.PictureBox pic1 
      Height          =   1215
      Left            =   240
      Picture         =   "frmthirdform.frx":2A0CB
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11183
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9960
      Width           =   975
   End
   Begin VB.CommandButton cmdMainmenu 
      BackColor       =   &H0000FFFF&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6706
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9840
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H0000FFFF&
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2303
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9960
      Width           =   975
   End
   Begin VB.Label lblShoot 
      BackColor       =   &H000000FF&
      Caption         =   "Shoot as many birds as possible!!! Good Luck!"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3623
      TabIndex        =   15
      Top             =   9240
      Width           =   7215
   End
End
Attribute VB_Name = "frmthirdform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Gun Selector (Zach Meyer's VB Project.vbp)
'Form Name : frmthirdform (frmthirdform.frm)
'Author: Zach Meyer
'Date Written: October 26, 2005
'Objective: This form is a game for the user to play that lets them shoot
                 'birds until they feel that they dont want to play anymore.
                 'Everytime a picture is clicked, birds on the form disappear,
                 'and then reappear when another pictre is clicked.
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

Dim player As String
Option Explicit

'This button will end the program.

Private Sub cmdExit_Click()
    End
End Sub

'This button will send the user to the gun finder form.

Private Sub cmdInfo_Click()
    frmfirstform.Hide
    frmsecondform.Show
    frmthirdform.Hide
End Sub

'This button will send the user to the main menu form.

Private Sub cmdMainmenu_Click()
    frmfirstform.Show
    frmsecondform.Hide
    frmthirdform.Hide
End Sub

'This will print the users name in lower left corner of the form,
'wishing them the best of luck during the game.

Private Sub Form_Activate()
    picName.Print "Good Luck " + player
 
End Sub

'This inputbox asks for the users name, once they select to play the game.

Private Sub Form_Load()
    player = InputBox("Enter Your Name", "Player Name")
    
End Sub

'This bird picture acts as a button, and it as well
'as the other pictures makes various of the other birds
'appear or dissappear depending on which bird picture is selected.

Private Sub pic1_Click()
    pic1.Visible = False
    pic2.Visible = True
    pic3.Visible = False
    pic4.Visible = True
    pic5.Visible = True
    pic6.Visible = True
    pic7.Visible = True
    pic8.Visible = True
    pic9.Visible = False
    pic10.Visible = True
    pic11.Visible = True
    pic12.Visible = True
End Sub

Private Sub pic10_Click()
    pic1.Visible = False
    pic2.Visible = True
    pic3.Visible = True
    pic4.Visible = True
    pic5.Visible = True
    pic6.Visible = True
    pic7.Visible = True
    pic8.Visible = False
    pic9.Visible = True
    pic10.Visible = False
    pic11.Visible = True
    pic12.Visible = True
End Sub

Private Sub pic11_Click()
    pic1.Visible = True
    pic2.Visible = True
    pic3.Visible = False
    pic4.Visible = True
    pic5.Visible = False
    pic6.Visible = True
    pic7.Visible = True
    pic8.Visible = False
    pic9.Visible = True
    pic10.Visible = True
    pic11.Visible = False
    pic12.Visible = True
End Sub

Private Sub pic12_Click()
    pic1.Visible = True
    pic2.Visible = True
    pic3.Visible = True
    pic4.Visible = True
    pic5.Visible = True
    pic6.Visible = True
    pic7.Visible = True
    pic8.Visible = True
    pic9.Visible = True
    pic10.Visible = False
    pic11.Visible = True
    pic12.Visible = False
End Sub

Private Sub pic2_Click()
    pic1.Visible = True
    pic2.Visible = False
    pic3.Visible = True
    pic4.Visible = True
    pic5.Visible = False
    pic6.Visible = True
    pic7.Visible = True
    pic8.Visible = True
    pic9.Visible = True
    pic10.Visible = True
    pic11.Visible = False
    pic12.Visible = True
End Sub

Private Sub pic3_Click()
    pic1.Visible = True
    pic2.Visible = True
    pic3.Visible = False
    pic4.Visible = True
    pic5.Visible = True
    pic6.Visible = False
    pic7.Visible = True
    pic8.Visible = True
    pic9.Visible = True
    pic10.Visible = True
    pic11.Visible = False
    pic12.Visible = True
End Sub

Private Sub pic4_Click()
    pic1.Visible = False
    pic2.Visible = False
    pic3.Visible = False
    pic4.Visible = False
    pic5.Visible = False
    pic6.Visible = True
    pic7.Visible = False
    pic8.Visible = True
    pic9.Visible = False
    pic10.Visible = False
    pic11.Visible = True
    pic12.Visible = False
End Sub

Private Sub pic5_Click()
    pic1.Visible = True
    pic2.Visible = True
    pic3.Visible = True
    pic4.Visible = True
    pic5.Visible = False
    pic6.Visible = True
    pic7.Visible = True
    pic8.Visible = True
    pic9.Visible = True
    pic10.Visible = True
    pic11.Visible = True
    pic12.Visible = True
End Sub

Private Sub pic6_Click()
    pic1.Visible = True
    pic2.Visible = True
    pic3.Visible = False
    pic4.Visible = True
    pic5.Visible = True
    pic6.Visible = False
    pic7.Visible = True
    pic8.Visible = False
    pic9.Visible = True
    pic10.Visible = True
    pic11.Visible = True
    pic12.Visible = False
End Sub

Private Sub pic7_Click()
    pic1.Visible = True
    pic2.Visible = False
    pic3.Visible = True
    pic4.Visible = True
    pic5.Visible = False
    pic6.Visible = True
    pic7.Visible = False
    pic8.Visible = True
    pic9.Visible = True
    pic10.Visible = False
    pic11.Visible = True
    pic12.Visible = True
End Sub

Private Sub pic8_Click()
    pic1.Visible = True
    pic2.Visible = True
    pic3.Visible = True
    pic4.Visible = True
    pic5.Visible = True
    pic6.Visible = True
    pic7.Visible = False
    pic8.Visible = False
    pic9.Visible = True
    pic10.Visible = True
    pic11.Visible = True
    pic12.Visible = True
End Sub

Private Sub pic9_Click()
    pic1.Visible = True
    pic2.Visible = False
    pic3.Visible = False
    pic4.Visible = False
    pic5.Visible = False
    pic6.Visible = False
    pic7.Visible = True
    pic8.Visible = True
    pic9.Visible = False
    pic10.Visible = True
    pic11.Visible = True
    pic12.Visible = False
End Sub

