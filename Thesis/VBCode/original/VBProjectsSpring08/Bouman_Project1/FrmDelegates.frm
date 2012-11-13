VERSION 5.00
Begin VB.Form FrmDelegates 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton CmdSearchDemo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search For Democratic Results By Specific State"
      DisabledPicture =   "FrmDelegates.frx":0000
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton CmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton CmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search For Republican Results By Specific State"
      DisabledPicture =   "FrmDelegates.frx":3794
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton CmdRepubDelegResults 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Republican Delegate Results As of 3/12/08"
      DisabledPicture =   "FrmDelegates.frx":AA25
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton CmdDisplayDemocraticDelegates 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Democratic Delegate Results As of 3/12/08"
      DisabledPicture =   "FrmDelegates.frx":11CB6
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton CmdTotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Delegate Count As of 3/12/08"
      DisabledPicture =   "FrmDelegates.frx":1544A
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox PicResults 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   4200
      ScaleHeight     =   7995
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label LblDelegateResults 
      BackColor       =   &H000000FF&
      Caption         =   "   Results By Delegates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FrmDelegates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DemoDelaState(1 To 1000) As String, ObamaDelegates(1 To 1000) As String, ClintonDelegates(1 To 1000) As String
Dim RepubDelaState(1 To 1000) As String, McCainDelegates(1 To 1000) As String, HuckabeeDelegates(1 To 1000) As String
Dim Pub As Integer, Ctr As Integer
    'Election Project
    'FrmPopVote
    'Ian Bouman
    'Written on 3/12
    'The purpose of this form is to load the arrays that hold the
    'results for the delegates and to display them for the user
    'to see. It also allows the user to search through the loaded
    'information for the results from a specific state.
    'The subroutines are similar to the ones on FrmPopVote.
    'It searches through the arrays for the results of a specific
    'state, and a specific political party, and displays that in the picture box.

Private Sub CmdDisplayDemocraticDelegates_Click()
Dim J As Integer
PicResults.Cls
For J = 1 To Ctr
    PicResults.Print DemoDelaState(J); Tab(20); ObamaDelegates(J); Tab(35); ClintonDelegates(J)
Next J
CmdSearchDemo.Enabled = True
End Sub


Private Sub CmdLoad_Click()
    'This command button loads the two arrays
Open App.Path & "/RepublicanDelegates.Txt" For Input As #1
Pub = 0
Ctr = 0
Do While Not EOF(1)
    Pub = Pub + 1
    Input #1, RepubDelaState(Pub), McCainDelegates(Pub)
Loop
Close
Open App.Path & "/DemocratDelegates.Txt" For Input As #1
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, DemoDelaState(Ctr), ObamaDelegates(Ctr), ClintonDelegates(Ctr)
Loop
Close
CmdTotal.Enabled = True
CmdDisplayDemocraticDelegates.Enabled = True
CmdRepubDelegResults.Enabled = True
End Sub

Private Sub CmdRepubDelegResults_Click()
    'This command button displays the results of the republican delegates
Dim J As Integer
PicResults.Cls
For J = 1 To Pub
    PicResults.Print RepubDelaState(J); Tab(20); McCainDelegates(J)
Next J
CmdSearch.Enabled = True
End Sub

Private Sub CmdReturn_Click()
FrmDelegates.Hide
FrmElection3.Show
End Sub

Private Sub CmdSearch_Click()
    'This command button searches through the republican results for those of a specific state
Dim Found As Boolean, FindMe As Integer, State As String
State = InputBox("Input the name of the state whose results you would like to see.", "Enter State")
FindMe = 0
Found = False
Do While Not Found And FindMe < Pub
    FindMe = FindMe + 1
    If State = RepubDelaState(FindMe) Then
        Found = True
        MsgBox ("The results for " & State & " are " & McCainDelegates(FindMe) & ".")
    End If
Loop
If Not Found Then
    MsgBox "Error: Misspelled state, or the listed state has not yet held its primaries.", , "Error"
End If
End Sub

Private Sub CmdSearchDemo_Click()
    'This searches through the Democrat results for delegates.
Dim Found As Boolean, FindMe As Integer, State As String
Found = False
FindMe = 0
State = InputBox("Input the name of the state whose results you would like to see.", "Enter State")
Do While Not Found And FindMe < Ctr
    FindMe = FindMe + 1
    If State = DemoDelaState(FindMe) Then
        Found = True
        MsgBox "The results for " & State & " are " & ObamaDelegates(FindMe) & " and " & ClintonDelegates(FindMe) & ".", , "Error"
    End If
Loop
If Not Found Then
    MsgBox ("Error: Misspelled state, or the listed state has not yet held its primaries.")
End If
End Sub

Private Sub CmdTotal_Click()
    'This button loads and prints the total delgates earned by the candidates
Dim PoliticalParty(1 To 10) As String, Pos As Integer, Name(1 To 10) As String, Delegates(1 To 10) As String
PicResults.Cls
Open App.Path & "/TotalDelegates.Txt" For Input As #1
Do While Not EOF(1)
    Pos = Pos + 1
    Input #1, PoliticalParty(Pos), Name(Pos), Delegates(Pos)
    PicResults.Print PoliticalParty(Pos); Tab(20); Name(Pos); Tab(40); Delegates(Pos)
Loop
Close
End Sub

