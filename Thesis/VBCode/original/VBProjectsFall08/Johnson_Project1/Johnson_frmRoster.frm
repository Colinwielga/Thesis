VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Johnson_frmRoster.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdBack 
      Caption         =   "BACK!"
      Height          =   1095
      Left            =   1440
      TabIndex        =   2
      Top             =   7800
      Width           =   3015
   End
   Begin VB.CommandButton cmdRoster 
      Caption         =   "Active Brewer 2008 Roster"
      Height          =   1095
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.PictureBox picRes 
      Height          =   9615
      Left            =   5040
      ScaleHeight     =   9555
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Milwaukee Brewers Fan Club Program 2008

'Form Name: Active Roster

'Author: Matthew Johnson

'Date Written: 11/2/2008

'Objective: In this form, I open a data file showing the Milwaukee Brewer's Active Roster.
'I thought it would be necessary to show the team members.  This is a program designed
'to educate the user on the Brewers.

Option Explicit
'Here I declare c (which is a counter), which can be used throughout the form under
'multiple sections; however, I don't use it in multiple sections.
Dim c As Integer

'This button allows me to go back to the initital form.
Private Sub CmdBack_Click()
    frmRoster.Hide
    frmIntro.Show
End Sub

Private Sub cmdRoster_Click()
'Here I assign the data into arrays
Dim number(1 To 40) As Integer, Player(1 To 40) As String, bt(1 To 40) As String, height(1 To 40) As String, weight(1 To 40) As Integer, birth(1 To 40) As String

picRes.Cls
picRes.Print "Number"; Tab(15); "Player"; Tab(40); "B/T"; Tab(55); "Height"; Tab(65); "Weight"; Tab(75); "Birth Date"
picRes.Print "*******************************************************************************************************************************************************************************************************************"
picRes.Print ""

c = 0

'Here I open data from a file (File Input)
Open App.Path & "\ActiveRoster.txt" For Input As #1

'Here it loops through the file until all the information from the file is printed.
    Do Until EOF(1)
        c = c + 1
        Input #1, number(c), Player(c), bt(c), height(c), weight(c), birth(c)
        picRes.Print number(c); Tab(15); Player(c); Tab(40); bt(c); Tab(55); height(c); Tab(65); weight(c); Tab(75); birth(c)
    Loop
Close #1
End Sub

