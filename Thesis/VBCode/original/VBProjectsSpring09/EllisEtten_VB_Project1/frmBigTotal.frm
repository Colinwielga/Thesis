VERSION 5.00
Begin VB.Form frmBigTotal 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate Total Points"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   4455
   End
   Begin VB.CommandButton cmdRestart 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Restart!"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   4455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   4455
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   5775
      Left            =   3120
      ScaleHeight     =   5715
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Points!"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "frmBigTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRestart_Click()
    frmBigTotal.Hide
    frmBigLittle.Show
    picResults
    Runningtotal = 0
End Sub

Private Sub cmdTotal_Click()
    picResults.Print "Service"; Tab(20); BServiceCTR
    picResults.Print "Social"; Tab(20); BSocialCTR
    picResults.Print "Meetings"; Tab(20); BMeetingCTR
    picResults.Print "*********************************************************"
    picResults.Print
    picResults.Print "Total Points"; Tab(20); Runningtotal
    picResults.Print
    Select Case Runningtotal
        Case Is > 200
            picResults.Print "Great job on exceeding the points needed! Keep it up girl!"
        Case 150 To 200
            picResults.Print "Just a few more events then you have 150! Let's do it!"
        Case 100 To 149
            picResults.Print "Lets go do some service and social events!"
        Case Is <= 99
            picResults.Print "Better attend some events. We gotta be at 200 points at the end of the semester!"
    End Select
End Sub
