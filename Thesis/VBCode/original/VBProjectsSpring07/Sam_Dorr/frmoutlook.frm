VERSION 5.00
Begin VB.Form frmoutlook 
   BackColor       =   &H000000C0&
   Caption         =   "2007 Team Outlook"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   Picture         =   "frmoutlook.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      MaskColor       =   &H000000C0&
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdsort 
      Caption         =   " Sort Team By Previous Ranking"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search For Team"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   2535
   End
   Begin VB.CommandButton cmdteams 
      Caption         =   "Show Top 25 Teams"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   2535
   End
   Begin VB.PictureBox picresults 
      Height          =   6615
      Left            =   4800
      ScaleHeight     =   6555
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmoutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'College World Series.(NCAACollegeWorldSeries.vbp)

'Form name: frmoutlook; Form caption: outlook

'Author: Sam Dorr

'Date written: March 25, 2007

' Form Objective: The objective of frmoutlook is to give an insight of potential
'                   teams for the 2007 College World Series.  frmoutlook places the
'               the top 30 teams, puts them in order by previous ranking and promotes
'                  a search function to find specific infromation about a team.

Option Explicit

Dim team(1 To 30) As String 'array
Dim points(1 To 30) As Integer 'array
Dim oldstanding(1 To 30) As Integer 'array
Dim ctr As Integer
Dim pos As Integer
Dim pos2 As Integer
Dim Found As Boolean
Dim teamname As String
Dim Pass As Integer
Dim I As Integer
Dim Temp As Integer
Dim Temp1 As Integer
Dim Temp2 As String





Private Sub cmdback_Click()
    frmoutlook.Hide
    frmhome.Show
End Sub
                                   
Private Sub cmdsearch_Click()

picresults.Cls

Found = False 'declares variable found as false

teamname = InputBox("Please enter the team you would like information for. (lower case)", "Team Name")

Do While ((Not Found) And (pos < ctr)) 'do loop until found=true and position is less than counter
    pos = pos + 1
    If teamname = team(pos) Then Found = True 'if position is found, found=true
Loop

If (Not Found) Then 'if position not found displays output
    picresults.Print "Sorry "; teamname; " was not a valid entry."
    
Else
    picresults.Print teamname; " has "; points(pos); " points and had a ranking of "; oldstanding(pos); "last month." 'if found, displays information about team
    
End If

End Sub

Private Sub cmdsort_Click()

picresults.Cls
For Pass = 1 To ctr - 1 'search function
    For pos2 = 1 To ctr - Pass 'for second postion to counter
        If oldstanding(pos2) > oldstanding(pos2 + 1) Then ' bubble sorts standings
            Temp = oldstanding(pos2)
            oldstanding(pos2) = oldstanding(pos2 + 1) '
            oldstanding(pos2 + 1) = Temp
            Temp1 = points(pos2) 'bubble sports corresponding points
            points(pos2) = points(pos2 + 1)
            points(pos2 + 1) = Temp1
            Temp2 = team(pos2) 'bubble sorts corresponding team
            team(pos2) = team(pos2 + 1)
            team(pos2 + 1) = Temp2
        End If
    Next pos2
Next Pass

picresults.Print "Standing Last Month"; Tab(25); "Team Name"; Tab(55); "Current Point Total"
picresults.Print "---------------------------------------------------------------------------------------------------------------------------"
    'displays the data from the search
For I = 1 To ctr
    picresults.Print oldstanding(I); Tab(25); team(I); Tab(55); points(I)
Next I

End Sub

Private Sub cmdteams_Click()

picresults.Cls 'clears picturebox

Open App.Path & "\teamdata.txt" For Input As #1 'opens text file

ctr = 0

Do Until EOF(1) 'puts data into array
    ctr = ctr + 1
    Input #1, team(ctr), points(ctr), oldstanding(ctr)
    picresults.Print team(ctr) 'only prints team name
Loop

Close #1

End Sub





