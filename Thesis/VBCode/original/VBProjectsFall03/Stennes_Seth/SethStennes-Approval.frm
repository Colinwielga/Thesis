VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   LinkTopic       =   "Form2"
   ScaleHeight     =   5955
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureDisplayBox 
      Height          =   3135
      Left            =   360
      ScaleHeight     =   3075
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   240
      Width           =   4695
   End
   Begin VB.PictureBox ElectionResults 
      Height          =   1815
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   3720
      Width           =   4695
   End
   Begin VB.CommandButton cmdComputeApproval 
      Caption         =   "Display The Election Results And The Winner When Using The Approval Voting System "
      Height          =   1455
      Left            =   5520
      TabIndex        =   2
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdGoToElectoralForm 
      Caption         =   "Go to another form to calculate the election results using The Electoral College"
      Height          =   1455
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdGoToBordaForm 
      Caption         =   "Go to another form to calculate the election results using The Borda Method"
      Height          =   1455
      Left            =   5520
      TabIndex        =   0
      Top             =   2160
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'VOTINGMETHODS(VOTINGMETHODS.vbp)
'Form1(SethStennes-Approval.frm)
'Seth Stennes
'October 28, 2003
'The purpose of this form is to compute the election results using the Approval Voting system.
'Then it is to display these results, state the winner, and display a picture of the winner.

'Computes and Displays Election Results, Winner, and picture of winner when using the Approval Voting System
Private Sub cmdComputeApproval_Click()
ElectionResults.Cls
PictureDisplayBox.Cls
ElectionResults.Print "Minnesota 2004 Presidential Election Results"
ElectionResults.Print "When Using the Approval Voting System"
ElectionResults.Print "****************************************************************************"
Dim Voters(1 To 12800) As String, I As Integer
Dim CandidateATotal As Integer, CandidateBTotal As Integer, CandidateCTotal As Integer, CandidateDTotal As Integer
Dim CandidateAName As String, CandidateBName As String, CandidateCName As String, CandidateDName As String
Open PATH & "approvalsystemdata.txt" For Input As #1
CandidateATotal = 0
CandidateBTotal = 0
CandidateCTotal = 0
CandidateDTotal = 0
For I = 1 To 12800
    Input #1, Voters(I)
    'ElectionResults.Print Voters(I)
    If Voters(I) = "CandidateA" Then
        CandidateATotal = CandidateATotal + 1
    End If
    If Voters(I) = "CandidateBCandidateC" Then
        CandidateBTotal = CandidateBTotal + 1
        CandidateCTotal = CandidateCTotal + 1
    End If
    If Voters(I) = "CandidateC" Then
        CandidateCTotal = CandidateCTotal + 1
    End If
    If Voters(I) = "CandidateDCandidateB" Then
        CandidateDTotal = CandidateDTotal + 1
        CandidateBTotal = CandidateBTotal + 1
    End If
Next I
CandidateAName = "Howard Dean"
CandidateBName = "George W. Bush"
CandidateCName = "Jesse Ventura"
CandidateDName = "Ralph Nader"
ElectionResults.Print CandidateAName, CandidateATotal; "Votes"
ElectionResults.Print CandidateBName, CandidateBTotal; "Votes"
ElectionResults.Print CandidateCName, CandidateCTotal; "Votes"
ElectionResults.Print CandidateDName, Tab(15), CandidateDTotal; "Votes"
ElectionResults.Print "****************************************************************************"
ElectionResults.Print CandidateCName; " Wins with"; CandidateCTotal; "votes"
PictureDisplayBox.Picture = LoadPicture("PATH & JesseVentura.jpg")
Close #1
End Sub

'Takes you to another location to compute and display results using the Borda Method
Private Sub cmdGoToBordaForm_Click()
Form2.Hide
Form3.Show
End Sub

'Takes you to another location to compute and display results using the Electoral College System
Private Sub cmdGoToElectoralForm_Click()
Form2.Hide
Form1.Show
End Sub
