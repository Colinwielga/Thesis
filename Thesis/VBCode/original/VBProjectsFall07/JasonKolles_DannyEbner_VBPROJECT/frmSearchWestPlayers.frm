VERSION 5.00
Begin VB.Form frmSearchWestPlayers 
   BackColor       =   &H000000FF&
   Caption         =   "Search for West Starters"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9360
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdBackHome 
      BackColor       =   &H00FF0000&
      Caption         =   "Click to Go to Home Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   8415
      Left            =   1320
      Picture         =   "frmSearchWestPlayers.frx":0000
      ScaleHeight     =   8355
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdload 
         BackColor       =   &H00FF0000&
         Caption         =   "Load the Players From a File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox picresults 
         Height          =   735
         Left            =   1440
         ScaleHeight     =   675
         ScaleWidth      =   4875
         TabIndex        =   4
         Top             =   5760
         Width           =   4935
      End
      Begin VB.CommandButton cmdInputSearchWestStarters 
         BackColor       =   &H00FF0000&
         Caption         =   "Click Here to Search for Western Conference Starters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmSearchWestPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim player(1 To 75) As String, points(1 To 75) As Integer
Dim rebounds(1 To 75) As Integer, assists(1 To 75) As Integer
Dim ctr As Integer, I As Integer

Private Sub cmdBackHome_Click()
frmHome.Show
frmSearchWestPlayers.Hide

End Sub

Private Sub cmdInputSearchWestStarters_Click()
Dim starter As String
Dim Found As Boolean
Found = False
I = 0

'clear the picture box
picresults.Cls

'bring up an input box
starter = InputBox("Enter a Western Conference Starter's Full Name", "Which Player Would You Like To Find?")

'print heading
picresults.Print "Starter"; Tab(28); "Points"; Tab(42); "Rebounds"; Tab(58); "Assists"
picresults.Print
'need to search your file for a player
Do While ((Not Found) And (I < ctr))
    I = I + 1
    If starter = player(I) Then Found = True
Loop

If (Not Found) Then
    MsgBox "The player you are looking for is not a Western Conference Starter", , "Try Another Name"
      Else
        picresults.Print starter, Tab(30); points(I); Tab(46); rebounds(I); Tab(58); assists(I)
End If

End Sub

Private Sub cmdload_Click()
'open the data file
Open App.Path & "\AllWestNotes.txt" For Input As #20

'get the data
ctr = 0
    Do While Not EOF(20)
        ctr = ctr + 1
        Input #20, player(ctr), points(ctr), rebounds(ctr), assists(ctr)
    Loop

cmdload.Visible = False

End Sub

Private Sub cmdQuit_Click()
End

End Sub

Private Sub Command1_Click()

End Sub
