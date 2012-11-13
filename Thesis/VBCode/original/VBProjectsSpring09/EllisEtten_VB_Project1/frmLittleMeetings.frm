VERSION 5.00
Begin VB.Form frmLittleMeetings 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCommittee 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   8
      Top             =   3720
      Width           =   4455
   End
   Begin VB.CommandButton cmdToTotal 
      BackColor       =   &H0000C000&
      Caption         =   "Go to Totals!"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   3975
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H0000C000&
      Caption         =   "Total for both"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   3975
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   5055
      Left            =   6600
      ScaleHeight     =   4995
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton cmdCommitteeMtg 
      BackColor       =   &H0000C000&
      Caption         =   "Add Committee Meetings"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   4335
   End
   Begin VB.CommandButton cmdMtg 
      BackColor       =   &H0000C000&
      Caption         =   "Meeting Points"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C000&
      Caption         =   "Quit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   3855
   End
   Begin VB.Label lblCommittee 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Please endter what committee you are on"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Label lblBigMtg 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Time for Meeting Points"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   9615
   End
End
Attribute VB_Name = "frmLittleMeetings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmeetings As Integer, Totalmeetings As Integer
Dim meetings As Integer, comtmtg(1 To 4) As Integer, committees(1 To 4) As String, cmeetings As Integer

Private Sub cmdCommitteeMtg_Click()
    Dim ctr As Integer, Found As Boolean, i As Integer, Committee As String
    Open App.Path & "\committees.txt" For Input As #1
    i = 0
    ctr = 0
    Found = False
    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, committees(ctr), comtmtg(ctr)
    Loop
    
    Committee = txtCommittee.Text
    Do While ((Not Found) And (i < ctr))
        i = i + 1
        If Committee = committees(i) Then
            Found = True
        End If
    Loop
    If Not Found Then
        picResults.Print "This committee hasn't held a meeting."
        picResults.Print
    Else
        picResults.Print Committee; " as had"; comtmtg(i); "meetings."
        picResults.Print
    End If
    cmeetings = comtmtg(i) * 5
    
    cmdTotal.Enabled = True
    cmdToTotal.Enabled = True
    cmdQuit.Enabled = True
    Close #1
End Sub

Private Sub cmdMtg_Click()
Dim meetings As Integer
    meetings = InputBox("How many meetings have you attended?", "Meetings")
    If tmeetings > 10 Then
        MsgBox "That is an invald number of meetings. Try again with less.", "Error"
    Else
        tmeetings = meetings * 10
        picResults.Print "You have"; tmeetings; "points from Tuesday meetings."
        picResults.Print
    End If
    
    txtCommittee.Enabled = True
    cmdCommitteeMtg.Enabled = True
    cmdMtg.Enabled = False
    cmdQuit.Enabled = True
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdTotal_Click()
    LMeetingCTR = cmeetings + tmeetings
    Runningtotal = Runningtotal + LMeetingCTR
    picResults.Print "You have"; LMeetingCTR; "meeting points."
End Sub

Private Sub cmdToTotal_Click()
    frmLittleTotal.Show
    frmLittleMeetings.Hide
End Sub

