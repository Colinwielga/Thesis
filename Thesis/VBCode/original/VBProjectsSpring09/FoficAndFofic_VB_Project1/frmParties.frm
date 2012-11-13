VERSION 5.00
Begin VB.Form frmPartiesAndCandidates 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Parites and Candidates"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   Picture         =   "frmParties.frx":0000
   ScaleHeight     =   1056.172
   ScaleMode       =   0  'User
   ScaleWidth      =   1931.694
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      Height          =   735
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7560
      Width           =   2655
   End
   Begin VB.PictureBox picResultsOfCandidates 
      BackColor       =   &H00FFFFFF&
      Height          =   5895
      Left            =   3120
      Picture         =   "frmParties.frx":205B6
      ScaleHeight     =   5835
      ScaleWidth      =   4395
      TabIndex        =   12
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton cmdDNZ 
      Height          =   1095
      Left            =   8400
      Picture         =   "frmParties.frx":40B6C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdHKDU 
      Height          =   1095
      Left            =   8400
      Picture         =   "frmParties.frx":44B39
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdHSSNHI 
      Height          =   1095
      Left            =   8400
      MaskColor       =   &H00FFC0FF&
      Picture         =   "frmParties.frx":48D20
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdSDP 
      Height          =   1095
      Left            =   8400
      Picture         =   "frmParties.frx":4D8AF
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdSBIH 
      Height          =   975
      Left            =   8400
      Picture         =   "frmParties.frx":515DC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdBPS 
      Height          =   1095
      Left            =   600
      Picture         =   "frmParties.frx":55779
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdHSP 
      Height          =   1095
      Left            =   600
      Picture         =   "frmParties.frx":5A278
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSDA 
      Height          =   1095
      Left            =   600
      Picture         =   "frmParties.frx":5E3FB
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdHDZBH 
      Height          =   1095
      Left            =   600
      Picture         =   "frmParties.frx":62124
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdHDZ1990 
      Height          =   975
      Left            =   600
      Picture         =   "frmParties.frx":66C34
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back"
      Height          =   735
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   135
   End
End
Attribute VB_Name = "frmPartiesAndCandidates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
 'Local Election for Prozor-Rama
    'frmPartiesAndCandidates
    'Josipa and Mario Fofic
    'Written 03/15/09'
    'The purpose of this form is to load the input that hold the
    'list for the candidates nominated by each political party
    'and to display them in picture box by their position on the election list.
    'comand buttons bellow load and display the list of candidates
    
Option Explicit
Dim Ctr As Integer
Dim Candidates(1 To 100) As String

    




Private Sub cmdBPS_Click()
picResultsOfCandidates.Cls   'this line of code will clear picResultsOfCandidates picturebox


Open App.Path & "\BPSCandidates.txt" For Input As #1 '
Ctr = 0

picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)                      'this loop reads data from file
     Ctr = Ctr + 1
    Input #1, Candidates(Ctr)             'this line of code gets input from the file
    picResultsOfCandidates.Print Candidates(Ctr)
Loop
Close #1           'this line close the input #1

End Sub

Private Sub cmdDNZ_Click()
Open App.Path & "\DNZCandidates.txt" For Input As #1
picResultsOfCandidates.Cls
Ctr = 0

picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Candidates(Ctr)
    picResultsOfCandidates.Print Candidates(Ctr) 'this line display the results in a picturebox
Loop
Close #1
End Sub

Private Sub cmdHDZ1990_Click()
picResultsOfCandidates.Cls



Open App.Path & "\hdz1990Candidates.txt" For Input As #1
Ctr = 0



picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Candidates(Ctr)
    picResultsOfCandidates.Print Candidates(Ctr)
Loop
Close #1
End Sub





Private Sub cmdHDZBH_Click()
picResultsOfCandidates.Cls

picResultsOfCandidates.Cls

Open App.Path & "\HDZCandidates.txt" For Input As #1
Ctr = 0


picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Candidates(Ctr)
    picResultsOfCandidates.Print Candidates(Ctr)
Loop
Close #1
End Sub

Private Sub cmdHKDU_Click()

picResultsOfCandidates.Cls

Open App.Path & "\HKDUCandidates.txt" For Input As #1
Ctr = 0

picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Candidates(Ctr)
    picResultsOfCandidates.Print Candidates(Ctr)
Loop
Close #1
End Sub

Private Sub cmdHSP_Click()
picResultsOfCandidates.Cls
Open App.Path & "\HSPCandidates.txt" For Input As #1
Ctr = 0

picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Candidates(Ctr)
    picResultsOfCandidates.Print Candidates(Ctr)
Loop
Close #1
End Sub

Private Sub cmdHSSNHI_Click()
picResultsOfCandidates.Cls
Open App.Path & "\HSSNHICandidates.txt" For Input As #1
Ctr = 0

picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Candidates(Ctr)
    picResultsOfCandidates.Print Candidates(Ctr)
Loop
Close #1
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSBIH_Click()
picResultsOfCandidates.Cls
Open App.Path & "\SBIHCandidates.txt" For Input As #1
Ctr = 0

picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Candidates(Ctr)
    picResultsOfCandidates.Print Candidates(Ctr)
Loop
Close #1
End Sub

Private Sub cmdSDA_Click()
picResultsOfCandidates.Cls
Open App.Path & "\SDACandidates.txt" For Input As #1
Ctr = 0

picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Candidates(Ctr)
    picResultsOfCandidates.Print Candidates(Ctr)
Loop
Close #1
End Sub

Private Sub cmdSDP_Click()
picResultsOfCandidates.Cls
Open App.Path & "\SDPCandidates.txt" For Input As #1
Ctr = 0

picResultsOfCandidates.Print "List of candidates:"
picResultsOfCandidates.Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Candidates(Ctr)
    picResultsOfCandidates.Print Candidates(Ctr)
Loop
Close #1
End Sub

Private Sub cmdBack_Click()
frmPartiesAndCandidates.Hide
frmProject.Show

End Sub


Private Sub picPicture1_Click()

End Sub


Private Sub Form_Load()

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2



    
End Sub
