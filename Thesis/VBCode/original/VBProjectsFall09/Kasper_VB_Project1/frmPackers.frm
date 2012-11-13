VERSION 5.00
Begin VB.Form frmPackers 
   Caption         =   "Packers"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRtn 
      Caption         =   "Return to Teams"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080FF80&
      FillColor       =   &H0080FF80&
      Height          =   4335
      Left            =   4080
      ScaleHeight     =   4275
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton cmddefense 
      Caption         =   "Defense"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOffense 
      Caption         =   "Offense"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   6120
      Left            =   0
      Picture         =   "frmPackers.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8760
   End
End
Attribute VB_Name = "frmPackers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author: Brandon Kasper
'Written 10/19/2009
'This form prints the starting lineup for Offense and Defense for the user in a picture box

Private Sub cmdDefense_Click()
    Dim Pos As Integer 'declares Pos as integer
    Open App.Path & "\PackersD.txt" For Input As #1 'opens the file Vikings offense
    Ctr = 0 'sets the value of the counter to 0
    Do Until EOF(1) 'starts the looping and sets it to the end of file
        Ctr = Ctr + 1 'adds a running total
        Input #1, PDnumb(Ctr), PDplayers(Ctr), PDpos(Ctr)
    Loop
    Close #1 'closes the form
     picResults.Cls 'clears the picture box
     picResults.Print "# Name ", "    ", "Position" 'prints in the picture box as shown
    For Pos = 1 To Ctr 'sets the range for the pos
        picResults.Print PDnumb(Pos); PDplayers(Pos); Tab(30); PDpos(Pos) 'prints the results
    Next Pos 'closes the pos
End Sub

Private Sub cmdOffense_Click()
    Dim Pos As Integer
    Open App.Path & "\PackersO.txt" For Input As #1 'opens the file Vikings offense
    Ctr = 0 'sets the value of the counter to 0
    Do Until EOF(1) 'starts the looping and sets it to the end of file
        Ctr = Ctr + 1 'adds a running total
        Input #1, POnumb(Ctr), POplayers(Ctr), POpos(Ctr)
    Loop
    Close #1
     picResults.Cls
     picResults.Print "# Name ", "    ", "Position"
    For Pos = 1 To Ctr
        picResults.Print POnumb(Pos); POplayers(Pos); Tab(30); POpos(Pos)
    Next Pos
End Sub

Private Sub cmdRtn_Click()
    frmPackers.Hide 'hides form from user
    frmTeams.Show 'shows form for user
End Sub
