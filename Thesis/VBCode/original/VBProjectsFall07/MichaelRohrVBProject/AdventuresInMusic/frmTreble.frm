VERSION 5.00
Begin VB.Form frmTreble 
   Caption         =   "Learning the Treble Clef"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   Picture         =   "frmTreble.frx":0000
   ScaleHeight     =   8400
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Test Yourself"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back To Main Page"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
   Begin VB.PictureBox picTreble 
      AutoSize        =   -1  'True
      Height          =   3345
      Left            =   1200
      Picture         =   "frmTreble.frx":2BBACA
      ScaleHeight     =   3285
      ScaleWidth      =   8415
      TabIndex        =   1
      Top             =   3360
      Width           =   8475
   End
   Begin VB.Label lblBass2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   $"frmTreble.frx":315BA8
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   9975
   End
   Begin VB.Label lblTreble 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Learning the Bass Clef"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmTreble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is informative only, the purpose of it is to show the Treble Clef and to inform the user the purpose of it musically and letter names on the staff
'It also has the same functions as many other pages, my having a button that goes back to frmLessonMainPage and also a button to take a quiz on

Private Sub cmdBack_Click()     'This button changes the form to frmLessonMainPage
    frmTreble.Hide                  'this hides frmTreble
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
End Sub

Private Sub cmdTest_Click()     'This button changes the form to frmTreble2
    frmTreble.Hide                  'this hides frmTreble
    frmTreble2.Show                 'this make frmTreble2 visible
End Sub
