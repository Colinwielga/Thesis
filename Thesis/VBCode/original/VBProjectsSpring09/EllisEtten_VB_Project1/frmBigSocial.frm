VERSION 5.00
Begin VB.Form frmBigSocial 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C000&
      Caption         =   "Quit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   3375
   End
   Begin VB.CommandButton cmdMeeting 
      BackColor       =   &H0000C000&
      Caption         =   "Go find meeting points!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   3735
   End
   Begin VB.CommandButton cmdCalcSocial 
      BackColor       =   &H0000C000&
      Caption         =   "Calculate your Social Points"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   4215
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H0000C000&
      Caption         =   "List Social Events"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   3615
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   5055
      Left            =   5880
      ScaleHeight     =   4995
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label lblBigSocial 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Find Your Social Points!"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   9255
   End
End
Attribute VB_Name = "frmBigSocial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Social(1 To 25) As String, ctr As Integer
Private Sub cmdCalcSocial_Click()
Dim points As Integer
    points = InputBox("Enter the number of social events you have attended", "Events attened")
    If points > 7 Then
        MsgBox "That is an invailed amount of events", , "Error"
    Else
        BSocialCTR = points * 10
        Runningtotal = Runningtotal + BSocialCTR
        picResults.Print "You have"; BSocialCTR; "service points."
    End If
    
    
    cmdMeeting.Enabled = True
    cmdQuit.Enabled = True
End Sub

Private Sub cmdList_Click()
   Open App.Path & "\social.txt" For Input As #1
    ctr = 0
    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, Social(ctr)
        picResults.Print Social(ctr)
    Loop
    picResults.Print
    picResults.Print
    cmdList.Enabled = False
    Close #1
End Sub

Private Sub cmdMeeting_Click()
    frmBigMeetings.Show
    frmBigSocial.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
