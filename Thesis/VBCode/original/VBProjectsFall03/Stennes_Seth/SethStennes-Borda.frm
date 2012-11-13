VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H000000FF&
   Caption         =   "Form3"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form3"
   ScaleHeight     =   6360
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureDisplayBox 
      Height          =   3135
      Left            =   600
      ScaleHeight     =   3075
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   240
      Width           =   4695
   End
   Begin VB.PictureBox ElectionResults 
      Height          =   1815
      Left            =   600
      ScaleHeight     =   1755
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   3720
      Width           =   4695
   End
   Begin VB.CommandButton cmdGoToApprovalForm 
      Caption         =   "Go to another form to calculate the election results using The Approval Voting System"
      Height          =   1455
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdComputeBorda 
      Caption         =   "Display The Election Results And The Winner When Using The Borda Method"
      Height          =   1455
      Left            =   5760
      TabIndex        =   1
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdGoToElectoralForm 
      Caption         =   "Go to another form to calculate the election results using The Electoral College System"
      Height          =   1455
      Left            =   5760
      TabIndex        =   0
      Top             =   2160
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'VOTINGMETHODS(VOTINGMETHODS.vbp)
'Form1(SethStennes-Borda.frm)
'Seth Stennes
'October 28, 2003
'The purpose of this form is to compute the election results using the Borda Method.
'Then it is to display these results, state the winner, and display a picture of the winner.

'Takes you to another location to compute and display results using the Approval Voting System
Private Sub cmdGoToApprovalForm_Click()
Form3.Hide
Form2.Show
End Sub

'Takes you to another location to compute and display results using the Electoral College System
Private Sub cmdGoToElectoralForm_Click()
Form3.Hide
Form1.Show
End Sub

'Computes and Displays Election Results, Winner, and picture of winner when using the Borda Method
Private Sub cmdComputeBorda_Click()
ElectionResults.Cls
PictureDisplayBox.Cls
ElectionResults.Print "Minnesota 2004 Presidential Election Results"
ElectionResults.Print "When Using the Borda Method"
ElectionResults.Print "****************************************************************************"
Dim Voters(1 To 12800) As String, I As Integer
Dim CandidateATotal As Integer, CandidateBTotal As Integer, CandidateCTotal As Integer, CandidateDTotal As Integer
Dim CandidateAName As String, CandidateBName As String, CandidateCName As String, CandidateDName As String
Open PATH & "bordamethoddata.txt" For Input As #1
CandidateATotal = 0
CandidateBTotal = 0
CandidateCTotal = 0
CandidateDTotal = 0
For I = 1 To 12800
    Input #1, Voters(I)
    'ElectionResults.Print Voters(I)
    If Voters(I) = "CandidateACandidateBCandidateC" Then
        CandidateATotal = CandidateATotal + 3
        CandidateBTotal = CandidateBTotal + 2
        CandidateCTotal = CandidateCTotal + 1
    End If
    If Voters(I) = "CandidateBCandidateACandidateC" Then
        CandidateBTotal = CandidateBTotal + 3
        CandidateATotal = CandidateATotal + 2
        CandidateCTotal = CandidateCTotal + 1
    End If
    If Voters(I) = "CandidateCCandidateBCandidateA" Then
        CandidateCTotal = CandidateCTotal + 3
        CandidateBTotal = CandidateBTotal + 2
        CandidateATotal = CandidateATotal + 1
    End If
    If Voters(I) = "CandidateDCandidateBCandidateC" Then
        CandidateDTotal = CandidateDTotal + 3
        CandidateBTotal = CandidateBTotal + 2
        CandidateCTotal = CandidateCTotal + 1
    End If
Next I
CandidateAName = "Howard Dean"
CandidateBName = "George W. Bush"
CandidateCName = "Jesse Ventura"
CandidateDName = "Ralph Nader"
ElectionResults.Print CandidateAName, CandidateATotal; "Points"
ElectionResults.Print CandidateBName, CandidateBTotal; "Points"
ElectionResults.Print CandidateCName, CandidateCTotal; "Points"
ElectionResults.Print CandidateDName, Tab(15), CandidateDTotal; "Points"
ElectionResults.Print "****************************************************************************"
ElectionResults.Print CandidateBName; " Wins with"; CandidateBTotal; "Points"
PictureDisplayBox.Picture = LoadPicture("PATH & GeorgeBush.jpg")
Close #1
End Sub
