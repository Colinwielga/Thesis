VERSION 5.00
Begin VB.Form frm1welcome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Welcome to Deal or No Deal"
   ClientHeight    =   8235
   ClientLeft      =   2775
   ClientTop       =   1020
   ClientWidth     =   10590
   FillColor       =   &H00800000&
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10590
   Begin VB.OptionButton optNo 
      Caption         =   "No, I need instructions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   5
      Top             =   3960
      Width           =   2415
   End
   Begin VB.OptionButton OptYes 
      Caption         =   "Yes Lets Play!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "I would like to QUIT"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   3
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmdno 
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   2
      Top             =   4800
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   7575
      Left            =   240
      Picture         =   "frm1welcome.frx":0000
      ScaleHeight     =   7515
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label lbltheme 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "<-- Click here to hear the Deal or No Deal theme song"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   7920
      TabIndex        =   7
      Top             =   6240
      Width           =   1815
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   855
      Left            =   6600
      OleObjectBlob   =   "frm1welcome.frx":6920
      SourceDoc       =   "M:\CS130\Project\theme.mp3"
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to Deal or No Deal. Have you played this game before?"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   6360
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "frm1welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project: Deal or No Deal
'frm1welcome
'Holly Reinking and Danielle Karp
'Written 3/15/09
'Purpose: To welcome the user and direct them towards instructions or the game

Private Sub cmdno_Click()           'Create an option button so players can choose whether or not they have played Deal or No Deal before

Sum = 3418416.01
frmmoney.cmdBanker.Enabled = False

If OptYes.Value = True Then             'We adapted this code from the online code offered at www.profsr.com/vb/vbless04.htm. It helped us discover how to write the code
    optNo.Value = False                 'If No, they are taken to the directions
    frmUsername.Show
    frm1welcome.Hide
ElseIf optNo.Value = True Then          'If Yes, they are taken to the frmUsername
    OptYes.Value = False
    frminstructions.Show
    frm1welcome.Hide
End If

End Sub
    
Private Sub cmdQuit_Click()             ' To QUIT the program
    End
End Sub


Private Sub Form_Load()                 'To load the option button
    OptYes.Value = False
    optNo.Value = False
End Sub



Private Sub picresults_Click()
    picResults.Picture = LoadPicture(App.Path & "\Howie.JPG")                                       'to load the picture of Deal or No Deal's Host Howie Mandel
        'http://media.photobucket.com/image/howie%2Bmandel/clyde67890/HowieMandel.jpg
        'The website above is where the picture of Howie Mandel (the host of Deal or No Deal) is found
End Sub

