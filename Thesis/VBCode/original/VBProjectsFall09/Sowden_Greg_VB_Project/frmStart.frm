VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Football"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20940
   LinkTopic       =   "Form1"
   Picture         =   "frmStart.frx":0000
   ScaleHeight     =   10140
   ScaleWidth      =   20940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit2 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H000000FF&
      Caption         =   "Begin Program"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   3135
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   10095
      Left            =   12600
      ScaleHeight     =   10035
      ScaleWidth      =   8355
      TabIndex        =   2
      Top             =   0
      Width           =   8415
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H000000FF&
      Caption         =   "Read The Data File"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   3135
   End
   Begin VB.Label lblStart 
      BackStyle       =   0  'Transparent
      Caption         =   "Football: The Offense"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

            


    


Option Explicit
'   Football: The Offense
'   Start
'   Greg Sowden
'   10/18/09
'   This form is the starting form of the project.
'   it has a button to read the data into the program right away.
'   Moving into the main form is disabled right away.
'   I did this because there are so many commands that require the data that the data needed to be read in right away

Private Sub cmdEnter_Click()
    frmStart.Hide
    frmRoster.Show
End Sub

Private Sub cmdQuit2_Click()
    End
End Sub

Private Sub cmdStart_Click()
'   this subcommand will read the data from the data file into 5 arrays
'   It will print the information in the picture box

    
    Open App.Path & "\Stats.txt" For Input As #1
                picResults1.Print Tab(0); "Name";
                picResults1.Print Tab(22); "Height (in)";
                picResults1.Print Tab(35); "Weight";
                picResults1.Print Tab(45); "Forty Time";
                picResults1.Print Tab(58); "Throwing Accuracy"
                
                picResults1.Print "***********************************************************************************"
'   a "do while" loop is required to read in the data
        Do While Not EOF(1)
            ctr = ctr + 1
            Input #1, names(ctr), heights(ctr), weights(ctr), forty(ctr), AccScore(ctr)
'   the 5 arrays are titled as such
                picResults1.Print Tab(0); names(ctr);
                picResults1.Print Tab(22); heights(ctr);
                picResults1.Print Tab(35); weights(ctr);
                picResults1.Print Tab(45); forty(ctr);
                picResults1.Print Tab(58); AccScore(ctr)
'   loop the statement to go through every input
        Loop
'   close the file once it has been read
            Close #1
'   enable the button to enter the main area of the program
            cmdStart.Enabled = True
            cmdEnter.Enabled = True
            
End Sub


Private Sub Form_Load()
'   disable enter until start is clicked
            cmdStart.Enabled = True
            cmdEnter.Enabled = False
End Sub
