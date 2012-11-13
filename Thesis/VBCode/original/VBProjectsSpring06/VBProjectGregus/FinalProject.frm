VERSION 5.00
Begin VB.Form frmSJULacrosse 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SJU Lacrosse"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Gill Sans Ultra Bold Condensed"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBoxCoverPic 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   5280
      Picture         =   "FINALP~1.frx":0000
      ScaleHeight     =   3135
      ScaleWidth      =   3375
      TabIndex        =   15
      Top             =   3600
      Width           =   3375
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "FINALP~1.frx":366C
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   14
      Top             =   2640
      Width           =   4455
   End
   Begin VB.CommandButton cmdBio 
      Caption         =   "Team Bio's"
      Height          =   855
      Left            =   7080
      TabIndex        =   13
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdFindStats 
      Caption         =   "Find Stats"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox picboxStats 
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   4635
      TabIndex        =   7
      Top             =   5160
      Width           =   4695
   End
   Begin VB.TextBox txtsearch 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Text            =   "Last name, First name"
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestion 
      Caption         =   "What's better than beating St. Thomas?"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   2535
   End
   Begin VB.PictureBox picboxLogo 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      Picture         =   "FINALP~1.frx":4797
      ScaleHeight     =   560.843
      ScaleMode       =   0  'User
      ScaleWidth      =   2236.208
      TabIndex        =   3
      Top             =   240
      Width           =   10365
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Quit"
      Height          =   975
      Left            =   9480
      TabIndex        =   2
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H80000004&
      Caption         =   "Team Roster"
      Height          =   855
      Left            =   7920
      MaskColor       =   &H00FFFF80&
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox picbox 
      Height          =   3255
      Left            =   9000
      ScaleHeight     =   3195
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblCredit 
      BackColor       =   &H00000000&
      Caption         =   "Project by: Dan Gregus"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label lblClick 
      Caption         =   "Click on the picture above to see the lighter side of SJU Lacrosse"
      Height          =   495
      Left            =   5280
      TabIndex        =   16
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Label lblPenalty 
      Caption         =   "Penalty Minutes"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblAssists 
      Caption         =   "Assists"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lblGoals 
      Caption         =   "Goals Scored"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Player Name"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblSearch 
      Caption         =   "Enter a player's name below to see his stats"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   4695
   End
End
Attribute VB_Name = "frmSJULacrosse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmBio (frmBio.frm)
'Dan Gregus
'3/19/06
'Objective: Provide a home page through which a user can easliy utilize several visual basic features to learn more about the lacrosse team as well as act as a hub that a user can use to access the rest of the project.
'Purpose of the Project: The purpose of this project is to create an interface through which a person unfamilliar with the Saint John's Lacrosse team can easily learn about a few key players on the roster.
    Dim playername As String
    Dim goals(1 To 14) As String
    Dim assists(1 To 14) As String
    Dim penalty(1 To 14) As String
    Dim pos As Integer
    Dim pass As Integer
    Dim I As Integer
    Dim tempteam As String
    Dim tempgoals As Integer
    Dim tempassists As Integer
    Dim temppenalty As Integer
    Dim roster(1 To 14) As String
    Dim temproster As String
    Dim size As Integer
    Dim rosterstats(1 To 14) As String


'Brings up the Biography page
Private Sub cmdBio_Click()
    frmSJULacrosse.Visible = False
    frmBio.Visible = True
End Sub

'Quits program
Private Sub cmdEnd_Click()
    End
End Sub


Private Sub cmdFindStats_Click()
picboxStats.Cls
pos = 0
Dim Found As Boolean
Found = False
'Displays player name, goals, assists, and penalty minutes in a picturebox
playername = txtsearch.Text
Do While ((Not Found) And (pos < size))
pos = pos + 1
If playername = rosterstats(pos) Then
    Found = True
    picboxStats.Print rosterstats(pos), Tab(20); goals(pos), Tab(35); assists(pos), Tab(45); penalty(pos)
End If
Loop

If (Not Found) Then
    picboxStats.Print ; "The name entered is not on the team"
End If
    picboxStats.Print
End Sub

Private Sub cmdOrder_Click()
    picboxStats.Cls
    
    'Loads and alphabetizes team roster into provided picture box.  Clears after every use.
    Open App.Path & "\roster.txt" For Input As #1
    pos = 0
    Do Until EOF(1)
        pos = pos + 1
        Input #1, roster(pos)
        Loop
        Close #1
    size = pos
    
    'Opens file rostersstats for use
    Open App.Path & "\rosterstats.txt" For Input As #2
    pos = 0
    Do Until EOF(2)
        pos = pos + 1
        Input #2, rosterstats(pos), goals(pos), assists(pos), penalty(pos)
        Loop
        Close #2
        
    For pass = 1 To size - 1
    For I = 1 To size - pass
        If roster(I) > roster(I + 1) Then
            temproster = roster(I)
            roster(I) = roster(I + 1)
            roster(I + 1) = temproster
            
        End If
        Next I
    Next pass
    For I = 1 To pos
        picbox.Print roster(I)
    Next I
End Sub

'Opens up frmNothing showcase the rivalry with St. Thomas
Private Sub cmdQuestion_Click()
    frmSJULacrosse.Visible = True
    frmNothing.Visible = True
End Sub

'link within a picture to the team picture page
Private Sub picBoxCoverPic_Click()
    frmSJULacrosse.Visible = False
    frmPicPage.Visible = True
End Sub

