VERSION 5.00
Begin VB.Form frmResults 
   BackColor       =   &H00FF0000&
   Caption         =   "Results"
   ClientHeight    =   7215
   ClientLeft      =   270
   ClientTop       =   660
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   7755
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Compare my Results"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Results"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5655
      Left            =   2400
      ScaleHeight     =   5595
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   840
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   240
      Picture         =   "frmResults.frx":0000
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label lblBy 
      BackColor       =   &H00FF0000&
      Caption         =   "Created by Jeff Doll"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label lbluser 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label lblresults 
      BackColor       =   &H00FF0000&
      Caption         =   "Results for"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCompare_Click()
    frmCompare.Show
End Sub
Private Sub cmdDisplay_Click()
'compute the score by adding all of the points together
score = Int(hundpts) + Int(LJpts) + Int(shotpts) + Int(highpts) + Int(quarterpts) + Int(HHpts) + Int(discpts) + Int(polepts) + Int(javpts) + Int(fifteenpts)
'print the outcome of all the events and the final score
picOutput.Print "Event", , "Mark", "Score"
picOutput.Print "*********************************************"
picOutput.Print "100 Meter Dash", onehund, Int(hundpts)
picOutput.Print "Long Jump", , LJ, Int(LJpts)
picOutput.Print "Shot Put", , shot, Int(shotpts)
picOutput.Print "High Jump", , high, Int(highpts)
picOutput.Print "400 Meter Dash", four, Int(quarterpts)
picOutput.Print "110 Meter High Hurdles", HH, Int(HHpts)
picOutput.Print "Discus Throw", , disc, Int(discpts)
picOutput.Print "Pole Vault", , pole, Int(polepts)
picOutput.Print "Javelin Throw", , jav, Int(javpts)
picOutput.Print "1500 Meter Run", fifteen, Int(fifteenpts)
picOutput.Print "**********************************************"
picOutput.Print "Your Score is", , , Int(score)
'if the qualifying mark was met then a message box will appear congratulating them
If Int(score) > 6000 Then
    MsgBox "CONGRATULATIONS! YOU HAVE MET THE DIVISION III PROVISIONAL QUALIFYING MARK OF 6000 POINTS!!", vbExclamation, "WOW!"
End If
End Sub
Private Sub cmdQuit_Click()
    'quit program
    End
End Sub
Private Sub Form_Load()
'users name will appear upon form loading
lbluser = n
End Sub
