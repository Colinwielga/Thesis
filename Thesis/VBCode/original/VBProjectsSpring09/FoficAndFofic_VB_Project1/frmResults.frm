VERSION 5.00
Begin VB.Form frmResults 
   BackColor       =   &H00404000&
   Caption         =   "Results"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   Picture         =   "frmResults.frx":0000
   ScaleHeight     =   8655
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   480
      Picture         =   "frmResults.frx":205B6
      ScaleHeight     =   1395
      ScaleWidth      =   7995
      TabIndex        =   6
      Top             =   120
      Width           =   8055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdCandAssembly 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Candidates in Assembly"
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.PictureBox picResultAssembly 
      Height          =   4815
      Left            =   3120
      Picture         =   "frmResults.frx":25993
      ScaleHeight     =   4755
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   1920
      Width           =   5415
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back"
      Height          =   1095
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearchCandidates 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Candidate Search"
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdResultsPolPart 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Results for Political Parties"
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Local Election for Prozor-Rama
    'frmResults
    'Josipa and Mario Fofic
    'Written 03/21/09
    'the purpose of this form is to direct the user to the results of
    'political parties or to the results of candidates
   
    


Private Sub cmdBack_Click()

frmResults.Hide
frmProject.Show
End Sub

Private Sub cmdCandAssembly_Click()
Dim Ctr As Integer
Dim Member(1 To 25) As String

picResultAssembly.Cls   'this line of code will clear picResultsOfCandidates picturebox


Open App.Path & "\CandidatesinAssembly.txt" For Input As #1 '
Ctr = 0

picResultAssembly.Print "List of candidates who became member of the Assembly:"
picResultAssembly.Print
Do While Not EOF(1)                      'this loop reads data from file
     Ctr = Ctr + 1
     Input #1, Member(Ctr)                'this line of code gets input from the file
     
    
    picResultAssembly.Print Member(Ctr)
Loop
Close #1           'this line close the input #1


End Sub


Private Sub cmdQuit_Click()
End
End Sub

'the purpose of this button is to search through the results of candidates
'it allows the user to search for specific candidate through inputbox,
'and diplay the results in messagebox




Private Sub cmdResultsCandidates_Click()


Private Sub cmdResultsPolPart_Click()
'This command button directs useer to the Results of the parties

frmResults.Hide
frmResultsParties.Show
End Sub


Private Sub cmdSearchCandidates_Click()
Dim Ctr As Integer
Dim Name(1 To 25) As String
Dim Party(1 To 25) As String
Dim I As Integer
Dim Name1 As String



Open App.Path & "\CandAssembly.txt" For Input As #1

I = 0
Ctr = 0
Found = False

Do While Not EOF(1)     'this line of code searches until found or end of list(match and stop searching)
    Ctr = Ctr + 1
    Input #1, Name(Ctr), Party(Ctr)
   
Loop
Close #1

Name1 = InputBox("Please, enter the last and first name you wish to find", "Name")
Do While ((Not Found) And (I < Ctr))

   I = I + 1
   If Name1 = Name(I) Then
       Found = True
   End If
Loop

If (Not Found) Then    'this line prints the results
  MsgBox "Sorry, " & Name1 & " is not the member of the Assembly.", , "Note"
 Else
  MsgBox "Candidate " & Name1 & " is the member of the Assembly.", , "Note"
End If

End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
  Left = Screen.Width / 2 - Width / 2
End Sub

