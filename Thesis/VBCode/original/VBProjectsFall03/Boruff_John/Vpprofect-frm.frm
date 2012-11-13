VERSION 5.00
Begin VB.Form MainForm1 
   BackColor       =   &H00C0C000&
   Caption         =   "Wax Project by John Boruff"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdWhy 
      BackColor       =   &H00C0C000&
      Caption         =   "Why Should You Wax?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9000
      Width           =   3255
   End
   Begin VB.CommandButton cmdWaxtype 
      BackColor       =   &H00C0C000&
      Caption         =   "What is the right wax for the snow Conditions?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9000
      Width           =   3375
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H00C0C000&
      Caption         =   "How much to spend on wax?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9000
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9000
      Width           =   3375
   End
   Begin VB.PictureBox picNordic1 
      Height          =   6615
      Left            =   9600
      Picture         =   "Vpprofect-frm.frx":0000
      ScaleHeight     =   6555
      ScaleWidth      =   5595
      TabIndex        =   3
      Top             =   0
      Width           =   5655
   End
   Begin VB.PictureBox picSnowboard1 
      Height          =   6615
      Left            =   5040
      Picture         =   "Vpprofect-frm.frx":1E44A
      ScaleHeight     =   6555
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   0
      Width           =   5055
   End
   Begin VB.PictureBox picSki1 
      Height          =   6615
      Left            =   0
      Picture         =   "Vpprofect-frm.frx":29A71
      ScaleHeight     =   6555
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.PictureBox Picture2 
         Height          =   3135
         Left            =   4800
         ScaleHeight     =   3135
         ScaleWidth      =   15
         TabIndex        =   1
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   $"Vpprofect-frm.frx":34744
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1680
      TabIndex        =   4
      Top             =   6720
      Width           =   11535
   End
End
Attribute VB_Name = "MainForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : WaxProject (VB-project.vbp)
'Form Name : MainForm1(Vbpofect-frm.frm)
'Author: John Boruff
'Date : Monday October 27, 2003
'Purpose of the Project: To have the user interact with the program
                    'to decide how much they want to spend on wax
                    ' and what the snow condition they will be
                    'experiancing so the program can educate them
                    'on the using the right wax for the condition
'purpose of the form:  It is the starting blocks of the project from, the
                    ' MainForm1 the user can navigate to different forms to
                    'learn about the type of wax he/she should use.  It also
                    'informs the user through message box's that pop up.
                    
Private Sub cmdPrice_Click()
    Priceform.Show  'goes to the Price form to decide how much to spen on wax.
    MainForm1.Hide
End Sub

Private Sub cmdQuit_Click()
End 'alows users to exit the program
End Sub

Private Sub cmdWaxtype_Click()
    SnowForm.Show   'goes to the  Snowform to dictate what is the right wax for the user.
    MainForm1.Hide
End Sub

Private Sub cmdWhy_Click() 'brings up message box to explain to the user why waxing is important to skiing and snowboarding.
    MsgBox "Waxing is very important.  If you are a Nordic Skier than glide is the most important thing to you, because less glide = more work.  Waxing is also important to Down Hill Skiers.  As with nordic skiing your glide will increase with the the right wax.  This means you could go faster, or just not get stuck on the cat walk.  Snowboarders often overlook waxing, but it is actually very important to them.  With now poles to push them along gilde is crushall, stridding to gain speed; gilde is the only thing keeping them moving.  Wether in the pipe or just freeriding a great wax job will help boarders keep their momentum.", , "Why you should wax"
End Sub
