VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   Picture         =   "character form.frx":0000
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Go back to Main Page"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   16
      Top             =   9120
      Width           =   1935
   End
   Begin VB.PictureBox Picture8 
      Height          =   2415
      Left            =   11520
      Picture         =   "character form.frx":9D835
      ScaleHeight     =   2355
      ScaleWidth      =   1755
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
   End
   Begin VB.PictureBox Picture7 
      Height          =   2415
      Left            =   11640
      Picture         =   "character form.frx":ABBF7
      ScaleHeight     =   2355
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox Picture6 
      Height          =   2175
      Left            =   6600
      Picture         =   "character form.frx":BB4D9
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   11
      Top             =   8880
      Width           =   2175
   End
   Begin VB.PictureBox Picture5 
      Height          =   2175
      Left            =   9360
      Picture         =   "character form.frx":CA66B
      ScaleHeight     =   2115
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   8880
      Width           =   1935
   End
   Begin VB.PictureBox Picture4 
      Height          =   2175
      Left            =   4080
      Picture         =   "character form.frx":D7BA5
      ScaleHeight     =   2115
      ScaleWidth      =   1755
      TabIndex        =   9
      Top             =   8880
      Width           =   1815
   End
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   1920
      Picture         =   "character form.frx":E529B
      ScaleHeight     =   2475
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   5760
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   1800
      Picture         =   "character form.frx":F3C0D
      ScaleHeight     =   2115
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdmrhowell 
      Caption         =   "Mr. Howell"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   6
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdmrshowell 
      Caption         =   "Mrs. Howell"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      TabIndex        =   5
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdginger 
      Caption         =   "Ginger"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   4
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdmaryann 
      Caption         =   "Mary Ann"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdprofessor 
      Caption         =   "Professor"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   2
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdskipper 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Skipper"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton cmdgilligan 
      Caption         =   "Gilligan"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label txtoutput 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   4320
      TabIndex        =   15
      Top             =   1440
      Width           =   6705
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Get to Know the Characters!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   14
      Top             =   120
      Width           =   13455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Gilligan's Island
'Form name:  character form
'Author:  Emily Olson
'Date written:  March 19, 2008
'Form Objective: inform user about the characters on "Gilligan's Island

Private Sub cmdgilligan_Click()
'clear data
    txtoutput = ""
'load and display data from file
    Dim Gilliganbio As String
        Open App.Path & "\Gilliganbio.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Gilliganbio
            txtoutput = txtoutput + Gilliganbio
        Loop
        Close #1
End Sub

Private Sub cmdginger_Click()
'clear data
    txtoutput = ""
'load and display data from file
    Dim gingerbio As String
        Open App.Path & "\gingerbio.txt" For Input As #1
        Do Until EOF(1)
            Input #1, gingerbio
            txtoutput = txtoutput + gingerbio
        Loop
        Close #1
End Sub

Private Sub cmdmaryann_Click()
'clear data
    txtoutput = ""
'load and display data from file
    Dim MaryAnnbio As String
        Open App.Path & "\MaryAnnbio.txt" For Input As #1
        Do Until EOF(1)
            Input #1, MaryAnnbio
            txtoutput = txtoutput + MaryAnnbio
        Loop
        Close #1
End Sub

Private Sub cmdmrhowell_Click()
'clear data
    txtoutput = ""
'load and display data from file
    Dim mrhowellbio As String
        Open App.Path & "\mrhowellbio.txt" For Input As #1
        Do Until EOF(1)
            Input #1, mrhowellbio
            txtoutput = txtoutput + mrhowellbio
        Loop
        Close #1
End Sub

Private Sub cmdmrshowell_Click()
'clear data
    txtoutput = ""
'load and display data from file
    Dim Mrshowellbio As String
        Open App.Path & "\Mrshowellbio.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Mrshowellbio
            txtoutput = txtoutput + Mrshowellbio
        Loop
        Close #1
End Sub

Private Sub cmdprofessor_Click()
'clear data
    txtoutput = ""
'load and display data from file
    Dim Professorbio As String
        Open App.Path & "\Professorbio.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Professorbio
            txtoutput = txtoutput + Professorbio
        Loop
        Close #1
End Sub

Private Sub cmdskipper_Click()
'clear data
    txtoutput = ""
'load and display data from file
    Dim Skipperbio As String
        Open App.Path & "\skipperbio.txt" For Input As #1
        Do Until EOF(1)
            Input #1, Skipperbio
            txtoutput = txtoutput + Skipperbio
        Loop
        Close #1
End Sub

Private Sub Command1_Click()
'load main page
    Form2.Hide
    Form1.Show
End Sub

Private Sub Form_Load()
'so program starts with form 1 since i created this page first instead of the main page
    Form2.Hide
    Form1.Show
End Sub

