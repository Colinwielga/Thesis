VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000013&
   Caption         =   "Help - Contacts"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn1 
      Caption         =   "Return to Home Page"
      Height          =   735
      Left            =   2520
      TabIndex        =   11
      Top             =   7080
      Width           =   2775
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   1320
      ScaleHeight     =   3555
      ScaleWidth      =   5235
      TabIndex        =   10
      Top             =   3000
      Width           =   5295
   End
   Begin VB.OptionButton optBoz 
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      Top             =   6120
      Width           =   255
   End
   Begin VB.OptionButton optJim 
      Height          =   255
      Left            =   6960
      TabIndex        =   2
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton optImad 
      Caption         =   "Option2"
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   4200
      Width           =   255
   End
   Begin VB.OptionButton optBrent 
      Caption         =   "Option1"
      Height          =   255
      Left            =   6960
      TabIndex        =   0
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblMyName 
      Caption         =   "By Brent Mergen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   12
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Label llblHelpTitle 
      Caption         =   "Help / References"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   $"frmHelp.frx":0000
      Height          =   735
      Left            =   4080
      TabIndex        =   8
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Warren 'Boz' Bostrom ACCT Dept's Tax Man"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   7
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label lblJim 
      Caption         =   "Jim Schnepf CSCI Genius"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label lblImad 
      Caption         =   "Imad Rahal      CSCI Mastermind"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label lblBrent 
      Caption         =   "Brent Mergen       Tax Programmer"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E-Z Tazes (Brent's E-ZTax Form VBProject.vbp)
'Help - Contacts (frmHelp)
'Brent Timothy Mergen
'24 March 2006
'This form let's you access any information from four business reps, people that are familiar with what thsi program does.


Private Sub cmdReturn1_Click()
    frmFrontpage.Show 'brings you to a new form
    frmHelp.Hide 'hides the current form
    MsgBox "You now have your contact info. Hopefully, your reference man gets back to you with an answer soon!", , "Help"
End Sub

Private Sub Form_Load()
    optBrent = False 'option button is declared false
    optImad = False 'option button is declared false
    optJim = False 'option button is declared false
    optBoz = Falsev 'option button is declared false
End Sub

Private Sub optBoz_Click()
Dim Helparray(1 To 100) As String
   Dim pos As Integer
   picHelp.Cls 'clears previous input displayed in picbox
   Open App.Path & "\HelpBoz.txt" For Input As #1 'opens Boz's Array
   pos = 0
   Do Until EOF(1)
        pos = pos + 1 'adds 1 to count the next line
        Input #1, Helparray(pos)
        picHelp.Print Helparray(pos) 'displays info in picbox
    Loop
    Close #1
End Sub

Private Sub optBrent_Click()
   Dim Helparray(1 To 100) As String
   Dim pos As Integer
   picHelp.Cls 'clears previous input displayed in picbox
   Open App.Path & "\HelpBrent.txt" For Input As #1 'opens Brent's Array
   pos = 0
   Do Until EOF(1)
        pos = pos + 1 'adds 1 to count the next line
        Input #1, Helparray(pos)
        picHelp.Print Helparray(pos) 'displays info in picbox
    Loop
    Close #1
End Sub

Private Sub optImad_Click()
Dim Helparray(1 To 100) As String
   Dim pos As Integer
   picHelp.Cls 'clears previous input displayed in picbox
   Open App.Path & "\HelpImad.txt" For Input As #1 'opens Imad's Array
   pos = 0
   Do Until EOF(1)
        pos = pos + 1 'adds 1 to count the next line
        Input #1, Helparray(pos)
        picHelp.Print Helparray(pos) 'displays info in picbox
    Loop
    Close #1
End Sub

Private Sub optJim_Click()
Dim Helparray(1 To 100) As String
   Dim pos As Integer
   picHelp.Cls 'clears previous input displayed in picbox
   Open App.Path & "\HelpJim.txt" For Input As #1 'opens Jim's Array
   pos = 0
   Do Until EOF(1)
        pos = pos + 1 'adds 1 to count the next line
        Input #1, Helparray(pos)
        picHelp.Print Helparray(pos) 'displays info in picbox
    Loop
    Close #1
End Sub
