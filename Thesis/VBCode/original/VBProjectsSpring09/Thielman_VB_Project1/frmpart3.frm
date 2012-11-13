VERSION 5.00
Begin VB.Form frmpart3 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdpictures 
      BackColor       =   &H0000C0C0&
      Caption         =   "Go to Pictures"
      Height          =   1095
      Left            =   7200
      MaskColor       =   &H00008080&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdcontinue 
      BackColor       =   &H0000C0C0&
      Caption         =   "Fun Facts"
      Height          =   1095
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H0000C0C0&
      Caption         =   "Play Again?"
      Height          =   1095
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdassists 
      BackColor       =   &H0080FFFF&
      Caption         =   "Lead Assists"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton cmdrebounds 
      BackColor       =   &H0080FFFF&
      Caption         =   "Lead Rebounder"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton cmdpoints 
      BackColor       =   &H0080FFFF&
      Caption         =   "Lead Point Scorer"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   2655
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H0080FFFF&
      Height          =   5535
      Left            =   1800
      ScaleHeight     =   5475
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   1200
      Width           =   5655
   End
   Begin VB.CommandButton cmdscore 
      BackColor       =   &H0080FFFF&
      Caption         =   "Individual Stats"
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lblaverage 
      BackStyle       =   0  'Transparent
      Caption         =   "Based on average per game"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8880
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lbltitle3 
      BackStyle       =   0  'Transparent
      Caption         =   "Stats"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   24
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmpart3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Timberwolves basketball
'frmpart3
'nick thielman
'3/18
'on this form the user may see the timberwolves leaders in points, rebounds, and assists
'they are also given the choice to return to the first form and start the game over again
'They can also move onto the fourth form. the stats are displayed in the picture box.
'With the different leader button the leader will appear first with decending order
'under them.
Option Explicit
Dim ctr As Integer, player(1 To 20) As String, rebounds(1 To 20) As Single, assists(1 To 20) As Single, points(1 To 20) As Single
Dim pass As Integer, pos As Integer, p As Integer
Dim tempplayer As String, temppoints As Single, tempassists As Single, temprebounds As Single



Private Sub cmdpictures_Click()
'takes the user to the picture form
frmpart3.Hide
frmpart5.Show
End Sub

Private Sub cmdscore_Click()

'clears picresults
 picresults.Cls
 
 ctr = 0
 'this will load an array
Open App.Path & "\playerstats.txt" For Input As #1

'prints header information
picresults.Print "Player"; Tab(15); "Rebounds"; Tab(30); "Assists"; Tab(40); "Points"
picresults.Print "***************************************************************"
 
 'prints the information in the array
 Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, player(ctr), rebounds(ctr), assists(ctr), points(ctr)
    picresults.Print player(ctr); Tab(20); rebounds(ctr); Tab(30); assists(ctr); Tab(40); points(ctr)

Loop
Close #1

'allows the other commands to be used
 cmdrebounds.Enabled = True
  cmdassists.Enabled = True
   cmdpoints.Enabled = True

End Sub
Private Sub cmdpoints_Click()
'this sorts the players by point total


'clears picresults
 picresults.Cls
 
'sort the names, assists, boards, and points with the lead point scorer on top
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If points(pos) < points(pos + 1) Then
            temppoints = points(pos)
            points(pos) = points(pos + 1)
            points(pos + 1) = temppoints
            tempplayer = player(pos)
            player(pos) = player(pos + 1)
            player(pos + 1) = tempplayer
            tempassists = assists(pos)
            assists(pos) = assists(pos + 1)
            assists(pos + 1) = tempassists
            temprebounds = rebounds(pos)
            rebounds(pos) = rebounds(pos + 1)
            rebounds(pos + 1) = temprebounds
            
        End If
    Next pos
Next pass
 
'prints the headers
   
    picresults.Print "Points"; Tab(10); "Player"; Tab(30); "Assists"; Tab(40); "Rebounds"
    picresults.Print "***********************************************************"
    
'then print the list
    For p = 1 To ctr
             picresults.Print points(p); Tab(10); player(p); Tab(30); assists(p); Tab(40); rebounds(p)
    Next p

End Sub

Private Sub cmdrebounds_Click()
'this sorts the players by rebound total

'clears picresults
 picresults.Cls
 
'sort the names, assists, points, and rebounds with the lead rebounder  on top
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If rebounds(pos) < rebounds(pos + 1) Then
            temprebounds = rebounds(pos)
            rebounds(pos) = rebounds(pos + 1)
            rebounds(pos + 1) = temprebounds
            tempplayer = player(pos)
            player(pos) = player(pos + 1)
            player(pos + 1) = tempplayer
            tempassists = assists(pos)
            assists(pos) = assists(pos + 1)
            assists(pos + 1) = tempassists
            temppoints = points(pos)
            points(pos) = points(pos + 1)
            points(pos + 1) = temppoints
            
        End If
    Next pos
Next pass
 
'prints the list
   
    picresults.Print "Rebounds"; Tab(15); "Player"; Tab(30); "Assists"; Tab(40); "Points"
    picresults.Print "***********************************************************"
    
'then print the list
    For p = 1 To ctr
             picresults.Print rebounds(p); Tab(10); player(p); Tab(30); assists(p); Tab(40); points(p)
    Next p
End Sub
Private Sub cmdassists_Click()
'this sorts the players by assist total

'clears picresults
 picresults.Cls
 
'sort the names, assists, boards, and points with the lead assists on top
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If assists(pos) < assists(pos + 1) Then
            tempassists = assists(pos)
            assists(pos) = assists(pos + 1)
            assists(pos + 1) = tempassists
            tempplayer = player(pos)
            player(pos) = player(pos + 1)
            player(pos + 1) = tempplayer
            temprebounds = rebounds(pos)
            rebounds(pos) = rebounds(pos + 1)
            rebounds(pos + 1) = temprebounds
            temppoints = points(pos)
            points(pos) = points(pos + 1)
            points(pos + 1) = temppoints
            
        End If
    Next pos
Next pass
 
'prints the list
   
    picresults.Print "Assists"; Tab(15); "Player"; Tab(30); "Rebounds"; Tab(40); "Points"
    picresults.Print "***********************************************************"
    
'then print the list
    For p = 1 To ctr
             picresults.Print assists(p); Tab(10); player(p); Tab(30); rebounds(p); Tab(40); points(p)
    Next p
End Sub

Private Sub cmdnew_click()
'allows the user to play again by taking them to the first form and hiding the third
frmpart3.Hide
frmpart1.Show
End Sub

Private Sub cmdcontinue_Click()
'takes user to next form
'close current form opens next form
frmpart3.Hide
frmpart4.Show
End Sub

