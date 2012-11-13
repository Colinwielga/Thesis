VERSION 5.00
Begin VB.Form frmSeriesPitching 
   BackColor       =   &H00C00000&
   Caption         =   "World Series Pitching Stats"
   ClientHeight    =   6150
   ClientLeft      =   2055
   ClientTop       =   3090
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   10335
   Begin VB.CommandButton cmdHitting 
      Caption         =   "Click here to go to Team Series Hitting Stats"
      Height          =   855
      Left            =   7920
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Click here to go back the the World Series Mainpage"
      Height          =   855
      Left            =   7920
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Meaning"
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      Height          =   1455
      Left            =   7680
      ScaleHeight     =   1395
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtAbbreviation 
      Height          =   285
      Left            =   7680
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblSerach 
      Caption         =   "If you Don't know what the Column Headings Mean, enter any one that you want to find out below. (Please Use Capital Letters"
      Height          =   975
      Left            =   7680
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   5940
      Left            =   0
      Picture         =   "frmSeriesPitching.frx":0000
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmSeriesPitching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: 1987 World Series
'Form name: frmSeriesPitching
'Authors: Hans Paul and Cole Wuollet
'Date Written: Wednesday November 1, 2006
'Objective: To display the pitching statistics for the World Series,
            'as well as to let the user search for the definitions
            'of various abbreviations in those statistics
Option Explicit

Private Sub cmdBack_Click() 'Hides Pitching Form and Shows The Series Form
    frmSeriesPitching.Hide
    frmSeriesStats.Show
End Sub

Private Sub cmdFind_Click()             'This Routine Searches for the Entered abbreviation
    picresults.Cls                      'And displays their corresponing meanings in
    Abbreviation = txtAbbreviation.Text 'a picturebox
    Select Case Abbreviation
    Case Is = "W"
        picresults.Print Abbreviation; " stands for Wins"
    Case "L"
        picresults.Print Abbreviation; " stands for Losses"
    Case "GS"
        picresults.Print Abbreviation; " stands for Games Started"
    Case "CG"
        picresults.Print Abbreviation; " stands for Complete Games"
    Case "S"
        picresults.Print Abbreviation; " stands for Saves"
    Case "SH"
        picresults.Print Abbreviation; " stands for Shutouts"
    Case "IP"
        picresults.Print Abbreviation; " stands for Innings Pitched"
    Case "ERA"
        picresults.Print Abbreviation; " stands for Earned Run Average"
    Case "H"
        picresults.Print Abbreviation; " stands for Hits"
    Case "SO"
        picresults.Print Abbreviation; " stands For Strike Outs"
    Case "ER"
        picresults.Print Abbreviation; " stands for Earned Runs"
    Case "BB"
        picresults.Print Abbreviation; " stands for Base on Balls (Walks)"
    Case Else
        MsgBox "Your entry does not match the abbreviations in the table, or it is not capitalized, please enter a new abbreviation", , "Error!"
    End Select
End Sub


Private Sub cmdHitting_Click() 'Hides Pitching Form and Shows Pitching Form
    frmSeriesPitching.Hide
    frmSeriesHitting.Show
End Sub
