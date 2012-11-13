VERSION 5.00
Begin VB.Form frmLDAC 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Leadership Development and Assessment Course"
   ClientHeight    =   11745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18870
   LinkTopic       =   "Form1"
   ScaleHeight     =   11745
   ScaleWidth      =   18870
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLDAC3 
      BackColor       =   &H00FFFF00&
      Height          =   5055
      Left            =   12360
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   7
      Top             =   5280
      Width           =   6255
   End
   Begin VB.PictureBox picLDAC2 
      BackColor       =   &H00FFFF00&
      Height          =   3615
      Left            =   12360
      ScaleHeight     =   3555
      ScaleWidth      =   6195
      TabIndex        =   6
      Top             =   1320
      Width           =   6255
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00C0C000&
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   11955
      TabIndex        =   5
      Top             =   1320
      Width           =   12015
   End
   Begin VB.CommandButton cmdWhy 
      BackColor       =   &H0000C000&
      Caption         =   "Why I Went?"
      Height          =   735
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picLDAC1 
      BackColor       =   &H00FFFF00&
      Height          =   8895
      Left            =   240
      ScaleHeight     =   8835
      ScaleWidth      =   11955
      TabIndex        =   2
      Top             =   1320
      Width           =   12015
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000C000&
      Caption         =   "Back to Main Form"
      Height          =   855
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10440
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C000&
      Caption         =   "Exit Program"
      Height          =   855
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10440
      Width           =   1815
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Tim Johnson   3/20    To State Why I went to LDAC and to show a bit about it"
      Height          =   615
      Left            =   15720
      TabIndex        =   8
      Top             =   10680
      Width           =   2895
   End
   Begin VB.Label lblLDACTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Leadership Development and Assessment Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   11535
   End
End
Attribute VB_Name = "frmLDAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()     'Goes back to Main Form
frmMain.Show                    'Goes back to Main Form
frmLDAC.Hide
End Sub

Private Sub cmdQuit_Click()     'Ends program where you are
    End                         'Ends program where you are
End Sub

Private Sub cmdWhy_Click()      'Answers a simple question

picInfo.Print "Every Army ROTC Cadet who enters into the Advanced Course attends the Leader Development and Assessment Course. It's a five-week summer course to evaluate"; Tab(3); "and train all Army ROTC Cadets in such events as Basic Rifle Marksmanship, Land Navigation, Physical fitness tests, and many more team building exercises. This"; Tab(3); "course normally takes place between your junior and senior years of college, and is conducted at Fort Lewis, Washington. Students maintain their normal daily"; Tab(3); "schedule as they develop their leadership and military skills in the classroom and in the field with the Army ROTC."

End Sub


Private Sub Form_Load()         'Puts ups pictures to improve form appearance

picLDAC1.Picture = LoadPicture(App.Path & "\" & ldacpix(1))
picLDAC2.Picture = LoadPicture(App.Path & "\" & ldacpix(2))
picLDAC3.Picture = LoadPicture(App.Path & "\" & ldacpix(3))

End Sub
