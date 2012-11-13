VERSION 5.00
Begin VB.Form frm1 
   BackColor       =   &H000040C0&
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit1 
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdswitch 
      Caption         =   "Click to view stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00FF8080&
      Caption         =   "Click button to start program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   1920
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   7155
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   1800
      Width           =   5175
   End
   Begin VB.CommandButton Cmdpic 
      Caption         =   "Click for a team picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Click this Button to Display the Roster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      Caption         =   "Created by: Kyle Hinners"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Caption         =   "    St. John's Lacrosse Main Page"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   6735
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Project1 (project1.vbp)
'frm1 (form1.frm)
'Kyle Hinners
'03/13/04
'the purpose of this form is to be a starting point for the program and is to let the user determine what they want the program to do for them
'the purpose of this program is to evaluate stats of the players and to display them
'it is useful to see the team leaders and is easily updated through notepad
'created march 12 by Kyle Hinners
'comp sci 130 with Professor Miller


Private Sub cmdload_Click()
'This command will print the roster, including number, name, and position
picresults.Cls
picresults.Print ; "Number"; Tab(12); "Name"; Tab(30); "Position"
picresults.Print "*************************************************************"
For j = 1 To 32
    picresults.Print ; Tab(3); number(j); Tab(10); names(j); Tab(30); position(j)
Next j
End Sub

Private Sub Cmdpic_Click()
'This will allow the user to go to form 3 which is the picture page.
frm3.Show
'This command makes the first form hidden so that only the picture form will be present
frm1.Hide
End Sub



Private Sub cmdquit1_Click()
'This command ends the program
End
End Sub

Private Sub cmdstart_Click()
'These various commands make the other buttons on the page inoperable until the start button is clicked
Cmdpic.Enabled = True
cmdswitch.Enabled = True
cmdquit1.Enabled = True
cmdload.Enabled = True
ctr = 0
'This opens the file team.txt into an array
'path equals M:\CS130\VBproject
Open path & "team.txt" For Input As #1
'this loads the file into an array until the end of the file is reached
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, number(ctr), names(ctr), position(ctr), goals(ctr), assists(ctr)
Loop
Close (1)
End Sub

Private Sub cmdswitch_Click()
'this command allows the user to switch to the second form which shows the stats
frm2.Show
'this causes form 1 to be hidden while viewing form 2
frm1.Hide
End Sub

Private Sub Form_Load()
'this makes the buttons inopoerable until the cmdstart button is clicked
cmdstart.Enabled = True
Cmdpic.Enabled = False
cmdswitch.Enabled = False
cmdquit1.Enabled = False
cmdload.Enabled = False
End Sub
