VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureDisplayBox 
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3075
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   120
      Width           =   4695
   End
   Begin VB.PictureBox ElectionResults 
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1755
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   3600
      Width           =   4695
   End
   Begin VB.CommandButton cmdGoToBordaForm 
      Caption         =   "Go to Another Form to Calcute the Election Results Using the Borda Method"
      Height          =   1455
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdGoToApprovalForm 
      Caption         =   "Go to another form to calculate the election results using The Approval Voting System"
      Height          =   1455
      Left            =   5400
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdComputeElectoral 
      Caption         =   "Display The Election Results And The Winner When Using The Electoral College System"
      Height          =   1455
      Left            =   5400
      TabIndex        =   0
      Top             =   3960
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'VOTINGMETHODS(VOTINGMETHODS.vbp)
'Form1(SethStennes-Electoral.frm)
'Seth Stennes
'October 28, 2003
'The purpose of this form is to compute the election results using the Electoral College system.
'Then it is to display these results, state the winner, and display a picture of the winner.

'Takes you to another location to compute and display results using the Approval Voting System
Private Sub cmdGoToApprovalForm_Click()
Form1.Hide
Form2.Show
End Sub

'Takes you to another location to compute and display results using the Borda Method
Private Sub cmdGoToBordaForm_Click()
Form1.Hide
Form3.Show
End Sub

'Computes and Displays Election Results, Winner, and picture of winner when using the Electoral College System
Private Sub cmdComputeElectoral_Click()
ElectionResults.Cls
PictureDisplayBox.Cls
ElectionResults.Print "Minnesota 2004 Presidential Election Results"
ElectionResults.Print "When Using the Electoral College"
ElectionResults.Print "****************************************************************************"
Dim Voters(1 To 12800) As String, I As Integer
Dim CandidateATotal As Integer, CandidateBTotal As Integer, CandidateCTotal As Integer, CandidateDTotal As Integer
Dim CandidateAPercentage As Single, CandidateBPercentage As Single, CandidateCPercentage As Single, CandidateDPercentage As Single
Dim CandidateAName As String, CandidateBName As String, CandidateCName As String, CandidateDName As String
Open PATH & "electoralsystemdata.txt" For Input As #1
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
    If Voters(I) = "CandidateB" Then
        CandidateBTotal = CandidateBTotal + 1
    End If
    If Voters(I) = "CandidateC" Then
        CandidateCTotal = CandidateCTotal + 1
    End If
    If Voters(I) = "CandidateD" Then
        CandidateDTotal = CandidateDTotal + 1
    End If
Next I
CandidateAPercentage = CandidateATotal / (I - 1)
CandidateBPercentage = CandidateBTotal / (I - 1)
CandidateCPercentage = CandidateCTotal / (I - 1)
CandidateDPercentage = CandidateDTotal / (I - 1)
CandidateAName = "Howard Dean"
CandidateBName = "George W. Bush"
CandidateCName = "Jesse Ventura"
CandidateDName = "Ralph Nader"
ElectionResults.Print CandidateAName, CandidateATotal, FormatPercent(CandidateAPercentage)
ElectionResults.Print CandidateBName, CandidateBTotal, FormatPercent(CandidateBPercentage)
ElectionResults.Print CandidateCName, CandidateCTotal, FormatPercent(CandidateCPercentage)
ElectionResults.Print CandidateDName, Tab(15), CandidateDTotal, FormatPercent(CandidateDPercentage)
ElectionResults.Print "****************************************************************************"
ElectionResults.Print CandidateAName; " Wins with "; FormatPercent(CandidateAPercentage, 2); " of the vote"
PictureDisplayBox.Picture = LoadPicture("PATH & HowardDean.jpg")
Close #1
End Sub
