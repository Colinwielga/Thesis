VERSION 5.00
Begin VB.Form frmLessonMainPage 
   Caption         =   "Main Page"
   ClientHeight    =   9555
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   Picture         =   "frmLessonMainPage.frx":0000
   ScaleHeight     =   9555
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdResults 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Click Here to See Your Quiz Results"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdClick 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click Me!!!"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      ScaleHeight     =   435
      ScaleWidth      =   7515
      TabIndex        =   13
      Top             =   600
      Width           =   7575
   End
   Begin VB.PictureBox picBassClef 
      AutoSize        =   -1  'True
      Height          =   1485
      Left            =   8400
      Picture         =   "frmLessonMainPage.frx":2C3B7A
      ScaleHeight     =   1425
      ScaleWidth      =   1485
      TabIndex        =   12
      Top             =   4560
      Width           =   1545
   End
   Begin VB.PictureBox picPiano 
      AutoSize        =   -1  'True
      Height          =   1740
      Left            =   8400
      Picture         =   "frmLessonMainPage.frx":2CAB10
      ScaleHeight     =   1680
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   2040
      Width           =   2415
   End
   Begin VB.PictureBox picTrebleClef 
      AutoSize        =   -1  'True
      Height          =   2295
      Left            =   1800
      Picture         =   "frmLessonMainPage.frx":2D79D2
      ScaleHeight     =   2235
      ScaleWidth      =   1290
      TabIndex        =   10
      Top             =   4320
      Width           =   1350
   End
   Begin VB.PictureBox picMusicBasics 
      AutoSize        =   -1  'True
      Height          =   1305
      Left            =   1920
      Picture         =   "frmLessonMainPage.frx":2E1168
      ScaleHeight     =   1245
      ScaleWidth      =   1200
      TabIndex        =   9
      Top             =   2160
      Width           =   1260
   End
   Begin VB.PictureBox picMiniBeethoven 
      AutoSize        =   -1  'True
      Height          =   2865
      Left            =   7200
      Picture         =   "frmLessonMainPage.frx":2E5F7A
      ScaleHeight     =   2805
      ScaleWidth      =   2250
      TabIndex        =   8
      Top             =   6360
      Width           =   2310
   End
   Begin VB.CommandButton cmdComposers 
      BackColor       =   &H000000FF&
      Caption         =   "List of Composers"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdMusic 
      BackColor       =   &H0000FFFF&
      Caption         =   "Music Basics"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdPiano 
      BackColor       =   &H0080FF80&
      Caption         =   "Notes on the Piano"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdBass 
      BackColor       =   &H000080FF&
      Caption         =   "Notes of the Staff:  Bass Cleff"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdTreble 
      BackColor       =   &H00FF8080&
      Caption         =   "Notes of the Staff: Treble Clef"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Choose Your Exercise You Would Like to See"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   8895
   End
End
Attribute VB_Name = "frmLessonMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The purpose of frmLessonMainPage is to display all of the options that the program offers
'using command buttons the user can choose five different buttons to take them to five different subject areas.
'It also has a button to end the program and a button to see the results of the quizes taken by the user.

Private Sub cmdBack_Click()         'This button changes forms
    frmLessonMainPage.Visible = False   'this hides frmLessonMainPage
    frmOpeningPage.Visible = True       'this makes frmOpeningPage visible
End Sub

Private Sub cmdBass_Click()     'This button changes forms
    frmLessonMainPage.Hide          'this hides frmLessonMainPage
    frmBass.Show                    'this make frmBass visible
End Sub

Private Sub cmdClick_Click()                                               'This button displays a welcome note in the picture box picResults
    picResults1.Print "Welcome "; NameGiven; " to Adventures in Music!!!"      'using the Public variable NameGiven which was acquired in frmOpeningPage
End Sub

Private Sub cmdComposers_Click()    'This button changes forms
    frmLessonMainPage.Hide              'this hides frmLessonMainPage
    frmComposers.Show                   'this makes frmComposers visible
End Sub

Private Sub cmdMusic_Click()        'This button changes forms
    frmLessonMainPage.Hide              'this hides frmLessonMainPage
    frmMusic1.Show                      'this makes frmMusic1 visible
End Sub

Private Sub cmdPiano_Click()    'This button changes forms
    frmLessonMainPage.Hide          'this hides frmLessonMainPage
    frmPiano1.Show                  'this makes frmPiano1 visible
End Sub

Private Sub cmdQuit_Click()                                                         'This button ends the program
    MsgBox "Goodbye " & NameGiven & ". Have a Great Day!", , "Have a Great Day!"    'This displays a message before the program completely ends by displaying a Message Box
End                                                                                 'the message box also uses the Public variable NameGiven acquired in frmOpeningPage
End Sub

Private Sub cmdResults_Click()  'This button changes forms to frmResults
    frmLessonMainPage.Hide          'this hides frmLessonMainPage
    frmResults.Show                 'this makes frmResults visible
End Sub

Private Sub cmdTreble_Click()   'This button changes forms
    frmLessonMainPage.Hide          'this hides frmLessonMainPage
    frmTreble.Show                  'this makes frmTreble visible
End Sub

