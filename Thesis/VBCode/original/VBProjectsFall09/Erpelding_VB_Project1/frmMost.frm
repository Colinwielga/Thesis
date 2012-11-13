VERSION 5.00
Begin VB.Form frmMost 
   BackColor       =   &H80000007&
   Caption         =   "Most of Each Statistic Per Player"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H00000080&
      Caption         =   "Enter the Name of a Player to see His Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00000080&
      Caption         =   "Look at the Averages per Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00000080&
      Caption         =   "Clear Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdMostAssists 
      BackColor       =   &H00000080&
      Caption         =   "Click to see Who had the Most Assists"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton cmdMostDigs 
      BackColor       =   &H00000080&
      Caption         =   "Click to see Who had the Most Digs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton cmdMostAces 
      BackColor       =   &H00000080&
      Caption         =   "Click to see Who had the Most Aces"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton cmdMostKills 
      BackColor       =   &H00000080&
      Caption         =   "Click to see Who had the Most Kills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   5895
      Left            =   3480
      ScaleHeight     =   5835
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "frmMost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Clear the picture box
Private Sub cmdClear_Click()
picResults.Cls
End Sub

'This button will bring up a picture box in which a players name will be entered and if
'the player is on the team his statistics will be printed in picture box

Private Sub cmdInput_Click()
Dim Player As String, Found As Boolean, I As Integer

'setup an inputbox
Player = InputBox("Please Enter the Name of the Player whose Statistics You'd like to See.", "Player")
'Set found to false
Found = False

For I = 1 To Ctr
    If Player = Names(I) Then
        Found = True
        picResults.Print "The player "; Player; " had "; Aces(I); " aces, "; Kills(I); " kills, "; Digs(I); " digs and "
        picResults.Print Assists(I); " assists last year."
    End If
Next I

If (Not Found) Then
    MsgBox ("The player you have entered was not on last years team")
End If


End Sub

'This subroutine will search through the file and list peoples number of aces in descending order
Private Sub cmdMostAces_Click()
'Declare variables that will be used in this subroutine
Dim tempAces As Integer, tempNames As String, I As Integer, Pass As Integer, Pos As Integer
Dim tempKills As Integer, tempDigs As Integer, tempAssists As Integer
'Use bubble sort feature to put array of kills in descending order
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Aces(Pos) < Aces(Pos + 1) Then
            tempAces = Aces(Pos)
            Aces(Pos) = Aces(Pos + 1)
            Aces(Pos + 1) = tempAces
            tempNames = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = tempNames
            tempKills = Kills(Pos)
            Kills(Pos) = Kills(Pos + 1)
            Kills(Pos + 1) = tempKills
            tempDigs = Digs(Pos)
            Digs(Pos) = Digs(Pos + 1)
            Digs(Pos + 1) = tempDigs
            tempAssists = Assists(Pos)
            Assists(Pos) = Assists(Pos + 1)
            Assists(Pos + 1) = tempAssists
        End If
    Next Pos
Next Pass

'Print the now sorted list in the picture box
picResults.Print "Name", "Number of Aces"
picResults.Print "***********************************"

For I = 1 To Ctr
    picResults.Print Names(I), Aces(I)
Next I
End Sub

'This subroutine will search through the file and list peoples number of assists in descending order
Private Sub cmdMostAssists_Click()
'Declare variables that will be used in this subroutine
Dim tempAssists As Integer, tempNames As String, I As Integer, Pass As Integer, Pos As Integer
Dim tempKills As Integer, tempDigs As Integer, tempAces As Integer
'Use bubble sort feature to put array of kills in descending order
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Assists(Pos) < Assists(Pos + 1) Then
            tempAssists = Assists(Pos)
            Assists(Pos) = Assists(Pos + 1)
            Assists(Pos + 1) = tempAssists
            tempNames = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = tempNames
            tempKills = Kills(Pos)
            Kills(Pos) = Kills(Pos + 1)
            Kills(Pos + 1) = tempKills
            tempDigs = Digs(Pos)
            Digs(Pos) = Digs(Pos + 1)
            Digs(Pos + 1) = tempDigs
            tempAces = Aces(Pos)
            Aces(Pos) = Aces(Pos + 1)
            Aces(Pos + 1) = tempAces
            
            
        End If
    Next Pos
Next Pass

'Print the now sorted list in the picture box
picResults.Print "Name", "Number of Assists"
picResults.Print "***********************************"

For I = 1 To Ctr
    picResults.Print Names(I), Assists(I)
Next I
End Sub

'This subroutine will search through the file and list peoples number of digs in descending order
Private Sub cmdMostDigs_Click()
'Declare variables that will be used in this subroutine
Dim tempDigs As Integer, tempNames As String, I As Integer, Pass As Integer, Pos As Integer
Dim tempKills As Integer, tempAces As Integer, tempAssists As Integer
'Use bubble sort feature to put array of kills in descending order
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Digs(Pos) < Digs(Pos + 1) Then
            tempDigs = Digs(Pos)
            Digs(Pos) = Digs(Pos + 1)
            Digs(Pos + 1) = tempDigs
            tempNames = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = tempNames
            tempKills = Kills(Pos)
            Kills(Pos) = Kills(Pos + 1)
            Kills(Pos + 1) = tempKills
            tempAces = Aces(Pos)
            Aces(Pos) = Aces(Pos + 1)
            Aces(Pos + 1) = tempAces
            tempAssists = Assists(Pos)
            Assists(Pos) = Assists(Pos + 1)
            Assists(Pos + 1) = tempAssists
        End If
    Next Pos
Next Pass

'Print the now sorted list in the picture box
picResults.Print "Name", "Number of Digs"
picResults.Print "***********************************"

For I = 1 To Ctr
    picResults.Print Names(I), Digs(I)
Next I
End Sub

'This subroutine will search through the file and list peoples number of kills in descending order
Private Sub cmdMostKills_Click()
'Declare variables that will be used in this subroutine
Dim tempKills As Integer, tempNames As String, tempAces As Integer, tempAssists As Integer, tempDigs As Integer
Dim I As Integer, Pass As Integer, Pos As Integer

'Use bubble sort feature to put array of kills in descending order
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Kills(Pos) < Kills(Pos + 1) Then
            tempKills = Kills(Pos)
            Kills(Pos) = Kills(Pos + 1)
            Kills(Pos + 1) = tempKills
            tempNames = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = tempNames
            tempAces = Aces(Pos)
            Aces(Pos) = Aces(Pos + 1)
            Aces(Pos + 1) = tempAces
            tempAssists = Assists(Pos)
            Assists(Pos) = Assists(Pos + 1)
            Assists(Pos + 1) = tempAssists
            tempDigs = Digs(Pos)
            Digs(Pos) = Digs(Pos + 1)
            Digs(Pos + 1) = tempDigs
        End If
    Next Pos
Next Pass

'Print the now sorted list in the picture box
picResults.Print "Name", "Number of Kills"
picResults.Print "***********************************"

For I = 1 To Ctr
    picResults.Print Names(I), Kills(I)
Next I
            
    
End Sub
'Switch to next form
Private Sub cmdNext_Click()
frmAverages.Show
frmMost.Hide
End Sub


