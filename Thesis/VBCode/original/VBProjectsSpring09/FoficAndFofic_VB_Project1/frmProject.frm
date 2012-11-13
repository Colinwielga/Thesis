VERSION 5.00
Begin VB.Form frmProject 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   1680
      Picture         =   "frmProject.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   7515
      TabIndex        =   3
      Top             =   1320
      Width           =   7575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdResults 
      BackColor       =   &H0000FFFF&
      Caption         =   "Election Results"
      Height          =   855
      Left            =   4200
      MaskColor       =   &H80000014&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdParty 
      BackColor       =   &H0000FFFF&
      Caption         =   "List of Political Parties and Candidates"
      Height          =   855
      Left            =   960
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   2175
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Local Election for Prozor-Rama
    'frmProject
    'Josipa and Mario Fofic
    'Written 03/10/09
 
    'The purpose of this project is to present elections for municipality Prozor-Rama
    'The purpose of 'frmProject' is to give the user two different directions:
    'List of political parties and candidates and election results.
    
    
Private Sub cmdParty_Click()
'This command button directs user to the list of parties and candidates


frmProject.Hide
frmPartiesAndCandidates.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdResults_Click()
'This command button directs user to the results of Elections

frmProject.Hide
frmResults.Show
End Sub



Private Sub Form_Load()
 
  AutoRedraw = True
  ScaleMode = vbPixels
  
  Top = Screen.Height / 2 - Height / 2
  Left = Screen.Width / 2 - Width / 2   'center form on the screen
  

    
  Text = "Election for Municipality Prozor-Rama"
  
    With Font
        .Name = "Arial"
        .Bold = False
        .Size = 22
    End With
                              
    
    ForeColor = vbGray                    'this lines add shadow to text
    CurrentX = 10
    CurrentY = 10
    Print Text

    ForeColor = vbWhite
    CurrentX = 10 - 7
    CurrentY = 10 - 7
    Print Text
End Sub
