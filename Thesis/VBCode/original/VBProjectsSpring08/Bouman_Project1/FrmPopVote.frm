VERSION 5.00
Begin VB.Form FrmPopVote 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form2"
   ScaleHeight     =   9105
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdEnd 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1335
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton CmdReturn2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return To Main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton CmdReturn1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton CmdLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   4080
      ScaleHeight     =   8655
      ScaleWidth      =   5895
      TabIndex        =   5
      Top             =   0
      Width           =   5895
   End
   Begin VB.CommandButton CmdSearchRepubs 
      BackColor       =   &H000000FF&
      Caption         =   "Search Republican Results In Specific State"
      DisabledPicture =   "FrmPopVote.frx":0000
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
      Height          =   1215
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton CmdSearchDems 
      BackColor       =   &H00FF0000&
      Caption         =   "Search Democrats Results In Specific State"
      DisabledPicture =   "FrmPopVote.frx":7291
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
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton CmdRepubicanResults 
      BackColor       =   &H000000FF&
      Caption         =   "Republican Popular Vote Results As of 3/12/08"
      DisabledPicture =   "FrmPopVote.frx":AA25
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton CmdDemPopVote 
      BackColor       =   &H00FF0000&
      Caption         =   "Democratic Popular Vote Results As of 3/12/08"
      DisabledPicture =   "FrmPopVote.frx":11CB6
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label LblPopVote 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Popular Vote Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "FrmPopVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DemCtr As Integer, RepCtr As Integer
Dim DemState(1 To 1000) As String, ObamaVote(1 To 1000) As String, ClintonVote(1 To 1000) As String
Dim RepState(1 To 1000) As String, McCainVote(1 To 1000) As String
    'Election Project
    'FrmPopVote
    'Ian Bouman
    'Written on 3/12
    'The purpose of this form is to load, display, and search
    'through the results of the two political parties.
    'The form begins by loading the two arrays, and then
    'the user can display the results from the party of his choice,
    'and search through the results for the results from a specific
    'state.
    

Private Sub CmdDemPopVote_Click()
    'This button prints the percentage of the democratic votes that the candidates received.
Dim J As Integer
PicResults.Cls
PicResults.Print "State"; Tab(20); "Obama"; Tab(35); "Clinton"
PicResults.Print "***************************************************"
For J = 1 To DemCtr
    PicResults.Print DemState(J); Tab(20); FormatPercent(ObamaVote(J)); Tab(35); FormatPercent(ClintonVote(J))
Next J
CmdSearchDems.Enabled = True

End Sub


Private Sub CmdEnd_Click()
End
End Sub

Private Sub CmdLoad_Click()
    'This button loads the files into arrays.
Open App.Path & "\DemocratPopularVote.Txt" For Input As #1
DemCtr = 0
Do While Not EOF(1)
    DemCtr = DemCtr + 1
    Input #1, DemState(DemCtr), ObamaVote(DemCtr), ClintonVote(DemCtr)
Loop
Close
Open App.Path & "\RepublicanPopularVote.Txt" For Input As #2
RepCtr = 0
Do While Not EOF(2)
    RepCtr = RepCtr + 1
    Input #2, RepState(RepCtr), McCainVote(RepCtr)
Loop
Close
MsgBox ("The file has been loaded.")
CmdDemPopVote.Enabled = True
CmdRepubicanResults.Enabled = True

End Sub

Private Sub CmdRepubicanResults_Click()
    'This button prints the republican results.
Dim J As Integer
PicResults.Cls
PicResults.Print "State"; Tab(20); "McCain"
PicResults.Print "**********************************"
For J = 1 To RepCtr
    PicResults.Print RepState(J); Tab(20); FormatPercent(McCainVote(J))
Next J
CmdSearchRepubs.Enabled = True
End Sub

Private Sub CmdReturn1_Click()
FrmPopVote.Hide
FrmElection3.Show

End Sub

Private Sub CmdReturn2_Click()
FrmPopVote.Hide
FrmElection1.Show
End Sub

Private Sub CmdSearchDems_Click()
    'This button prints the democratic results that the user searched for.
Dim Found As Boolean, Pos As Integer, SearchState As String
Found = False
SearchState = InputBox("Input the state whose results you would like to see.", "Input State")
Do While Not Found And Pos < DemCtr
    Pos = Pos + 1
    If SearchState = DemState(Pos) Then
        Found = True
        MsgBox ("The results of the popular vote in " & SearchState & " is " & FormatPercent(ObamaVote(Pos)) & " to Barack Obama and " & FormatPercent(ClintonVote(Pos)) & " to Hillary Clinton.")
    End If
Loop
If Not Found Then
    MsgBox "Error: Misspelled state or the entered state has not had its primaries/caucus yet.", , "Error"
End If
End Sub

Private Sub CmdSearchRepubs_Click()
    'This button allows the user to search through the results for those of a specific state.
Dim Found As Boolean, Pos As Integer, State As String
Found = False
State = InputBox("Input the state whose results you would like to see.", "Input State")
Do While Not Found And Pos < RepCtr
    Pos = Pos + 1
    If State = RepState(Pos) Then
        Found = True
        MsgBox ("The results of the popular vote in " & State & " is " & FormatPercent(McCainVote(Pos)) & " to John McCain.")
    End If
Loop
If Not Found Then
    MsgBox "Error: Misspelled state or the entered state has not had its primaries/caucus yet.", , "Error"
End If
End Sub

