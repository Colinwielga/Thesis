VERSION 5.00
Begin VB.Form frmSeriesHitting 
   BackColor       =   &H00FF0000&
   Caption         =   "World Series Hitting Stats"
   ClientHeight    =   8655
   ClientLeft      =   2055
   ClientTop       =   3090
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10470
   Begin VB.CommandButton cmdSearch 
      Caption         =   "If you Don't know what the Column Headings Mean,Click here to enter any one that you want to find out. (Please Use Capital Letters"
      Height          =   975
      Left            =   7800
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.PictureBox picresults 
      Height          =   1455
      Left            =   7800
      ScaleHeight     =   1395
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Meaning"
      Height          =   495
      Left            =   7800
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Click here to go back the the World Series Mainpage"
      Height          =   855
      Left            =   7800
      TabIndex        =   1
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdpitching 
      Caption         =   "Click here to go to Team Series Pitching Stats"
      Height          =   855
      Left            =   7800
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   13980
      Left            =   0
      Picture         =   "frmSeriesHitting.frx":0000
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "frmSeriesHitting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: 1987 World Series
'Form name: frmSeriesHitting
'Authors: Hans Paul and Cole Wuollet
'Date Written: Wednesday November 2, 2006
'Objective: To display the hitting statistics for the World Series,
            'as well as to let the user search for the definitions
            'of various abbreviations in those statistics
Option Explicit

Private Sub cmdBack_Click() 'Hides the Hitting Form
    frmSeriesHitting.Hide
End Sub

Private Sub cmdFind_Click()             'This Routine Searches for the Entered abbreviation
    picResults.Cls                      'And displays their corresponing meanings in
                                        'a picturebox
    Select Case Abbreviation
    Case Is = "Pos"
        picResults.Print Abbreviation; " stands for Position Played"
    Case "G"
        picResults.Print Abbreviation; " stands for Games Played"
    Case "AB"
        picResults.Print Abbreviation; " stands for At Bats"
    Case "H"
        picResults.Print Abbreviation; " stands for Hits"
    Case "2B"
        picResults.Print Abbreviation; " stands for Doubles"
    Case "3B"
        picResults.Print Abbreviation; " stands for Triples"
    Case "HR"
        picResults.Print Abbreviation; " stands for Home Runs"
    Case "R"
        picResults.Print Abbreviation; " stands for Runs"
    Case "RBI"
        picResults.Print Abbreviation; " stands for Runs Batted In"
    Case "AVG"
        picResults.Print Abbreviation; " stands for Batting Average"
    Case "BB"
        picResults.Print Abbreviation; " stands for Base on Balls (Walks"
    Case "SO"
        picResults.Print Abbreviation; " stands for Strikeouts"
    Case "SB"
        picResults.Print Abbreviation; " stands for Stolen Bases"
    Case Else
        MsgBox "Your entry does not match an abbreviation in the table, or it is not capitalized, please enter a new abbreviation", , "Error!"
    End Select
End Sub

Private Sub cmdpitching_Click() 'Hides the Hitting Form and Shows the Pitching Form
    frmSeriesPitching.Show
    frmSeriesHitting.Hide
End Sub

Private Sub cmdSearch_Click()
Abbreviation = InputBox("Enter an Abbreviation", "Abbreviation")
End Sub
