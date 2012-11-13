VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   Picture         =   "form1togethert.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnextform 
      Caption         =   "NEXT PAGE"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2040
      TabIndex        =   11
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   195
      Left            =   2040
      TabIndex        =   10
      Top             =   7740
      Width           =   1695
   End
   Begin VB.CommandButton cmdjersey 
      Caption         =   "view roster arranged by jersey number"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdplusminus 
      Caption         =   "view roster arranged by plus minus"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdpenalty 
      Caption         =   "view roter arranged by total penalty minutes"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdgames 
      Caption         =   "view roster arranged by total games played"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdpoints 
      Caption         =   "view roster arranged by points"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdassists 
      Caption         =   "view roster arranged by assists"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdgoals 
      Caption         =   "view roster arranged by goals scored"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdindividual 
      Caption         =   "Individual Stats"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9720
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.PictureBox picresults 
      Height          =   6975
      Left            =   4200
      ScaleHeight     =   6915
      ScaleWidth      =   8115
      TabIndex        =   1
      Top             =   1440
      Width           =   8175
   End
   Begin VB.CommandButton cmdread 
      Caption         =   "view Wild stats for 2002-2003 season"
      Height          =   615
      Left            =   9720
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim names(1 To 100) As String
Dim games(1 To 100) As Integer
Dim goals(1 To 100) As Integer
Dim assists(1 To 100) As Integer
Dim points(1 To 100) As Integer
Dim penalty(1 To 100) As Integer
Dim plusminus(1 To 100) As Integer
Dim jersey(1 To 100) As Integer
Dim x As Integer
Dim n As String
Dim found As Boolean
Dim ctr As Integer
Dim pass As Integer
Dim comp As Integer
Dim tempgoals As Integer
Dim tempnames As String
Dim tempjersey As Integer
Dim tempgames As Integer
Dim tempassists As Integer
Dim temppoints As Integer
Dim temppenalty As Integer
Dim tempplusminus As Integer
Dim path As String














Private Sub cmdassists_Click()
'view roster arranged by assists
ctr = 32
picresults.Cls

For pass = 1 To ctr - 1
    For comp = 1 To ctr - 1
        If assists(comp) < assists(comp + 1) Then
        
        'switch assists
        tempassists = assists(comp)
        assists(comp) = assists(comp + 1)
        assists(comp + 1) = tempassists
        
        'switch goals
        tempgoals = goals(comp)
        goals(comp) = goals(comp + 1)
        goals(comp + 1) = tempgoals
        
        'switch names
        tempnames = names(comp)
        names(comp) = names(comp + 1)
        names(comp + 1) = tempnames
        
        'switch jersey
        tempjersey = jersey(comp)
        jersey(comp) = jersey(comp + 1)
        jersey(comp + 1) = tempjersey
        
        'switch games played
        tempgames = games(comp)
        games(comp) = games(comp + 1)
        games(comp + 1) = tempgames
        
        
        'switch points
        temppoints = points(comp)
        points(comp) = points(comp + 1)
        points(comp + 1) = temppoints
        
        'switch penalty
        temppenalty = penalty(comp)
        penalty(comp) = penalty(comp + 1)
        penalty(comp + 1) = temppenalty
        
        'switch plus minus
        tempplusminus = plusminus(comp)
        plusminus(comp) = plusminus(comp + 1)
        plusminus(comp + 1) = tempplusminus
        
        End If
        Next comp
        Next pass
        
     
picresults.Print ; "Player                            jersey#   Games-played   Goals       Assists       Points    Penalty Minutes      +/- "
picresults.Print
For x = 1 To 32
picresults.Print names(x);
picresults.Print Tab(25); jersey(x);
picresults.Print Tab(37); games(x); Tab(50); goals(x); Tab(60); assists(x); Tab(70); points(x); Tab(80); penalty(x); Tab(95); plusminus(x)
            
        Next x
        
        
End Sub


Private Sub cmdgames_Click()
'arrange by total games played
ctr = 32
picresults.Cls

For pass = 1 To ctr - 1
    For comp = 1 To ctr - 1
        If games(comp) < games(comp + 1) Then
        
        'switch goals
        tempgoals = goals(comp)
        goals(comp) = goals(comp + 1)
        goals(comp + 1) = tempgoals
        
        'switch names
        tempnames = names(comp)
        names(comp) = names(comp + 1)
        names(comp + 1) = tempnames
        
        'switch jersey
        tempjersey = jersey(comp)
        jersey(comp) = jersey(comp + 1)
        jersey(comp + 1) = tempjersey
        
        'switch games played
        tempgames = games(comp)
        games(comp) = games(comp + 1)
        games(comp + 1) = tempgames
        
        'switch assists
        tempassists = assists(comp)
        assists(comp) = assists(comp + 1)
        assists(comp + 1) = tempassists
        
        'switch points
        temppoints = points(comp)
        points(comp) = points(comp + 1)
        points(comp + 1) = temppoints
        
        'switch penalty
        temppenalty = penalty(comp)
        penalty(comp) = penalty(comp + 1)
        penalty(comp + 1) = temppenalty
        
        'switch plus minus
        tempplusminus = plusminus(comp)
        plusminus(comp) = plusminus(comp + 1)
        plusminus(comp + 1) = tempplusminus
        
        End If
        Next comp
        Next pass
        
     
picresults.Print ; "Player                            jersey#   Games-played   Goals       Assists       Points    Penalty Minutes      +/- "
picresults.Print
For x = 1 To 32
picresults.Print names(x);
picresults.Print Tab(25); jersey(x);
picresults.Print Tab(37); games(x); Tab(50); goals(x); Tab(60); assists(x); Tab(70); points(x); Tab(80); penalty(x); Tab(95); plusminus(x)
            
        Next x
        
End Sub

Private Sub cmdgoals_Click()
'to arrange the roster by goals scored, descending order
ctr = 32
picresults.Cls

For pass = 1 To ctr - 1
    For comp = 1 To ctr - 1
        If goals(comp) < goals(comp + 1) Then
        
        'switch goals
        tempgoals = goals(comp)
        goals(comp) = goals(comp + 1)
        goals(comp + 1) = tempgoals
        
        'switch names
        tempnames = names(comp)
        names(comp) = names(comp + 1)
        names(comp + 1) = tempnames
        
        'switch jersey
        tempjersey = jersey(comp)
        jersey(comp) = jersey(comp + 1)
        jersey(comp + 1) = tempjersey
        
        'switch games played
        tempgames = games(comp)
        games(comp) = games(comp + 1)
        games(comp + 1) = tempgames
        
        'switch assists
        tempassists = assists(comp)
        assists(comp) = assists(comp + 1)
        assists(comp + 1) = tempassists
        
        'switch points
        temppoints = points(comp)
        points(comp) = points(comp + 1)
        points(comp + 1) = temppoints
        
        'switch penalty
        temppenalty = penalty(comp)
        penalty(comp) = penalty(comp + 1)
        penalty(comp + 1) = temppenalty
        
        'switch plus minus
        tempplusminus = plusminus(comp)
        plusminus(comp) = plusminus(comp + 1)
        plusminus(comp + 1) = tempplusminus
        
        End If
        Next comp
        Next pass
        
     
picresults.Print ; "Player                            jersey#   Games-played   Goals       Assists       Points    Penalty Minutes      +/- "
picresults.Print
For x = 1 To 32
picresults.Print names(x);
picresults.Print Tab(25); jersey(x);
picresults.Print Tab(37); games(x); Tab(50); goals(x); Tab(60); assists(x); Tab(70); points(x); Tab(80); penalty(x); Tab(95); plusminus(x)
            
        Next x
        
        
End Sub

Private Sub cmdindividual_Click()
'get a name from user and show individual stats
n = InputBox("Enter the name of a Wild Player from the roster")
picresults.Cls
found = False
For x = 1 To 32
' if name entered in list prints individual stats
    If n = names(x) Then
    found = True
    picresults.Print ; "Player                            jersey#   Games-played   Goals       Assists       Points    Penalty Minutes      +/- "
    picresults.Print
    picresults.Print names(x);
    picresults.Print Tab(25); jersey(x);
    picresults.Print Tab(37); games(x); Tab(50); goals(x); Tab(60); assists(x); Tab(70); points(x); Tab(80); penalty(x); Tab(95); plusminus(x)

    End If
    Next x
    If found = False Then
    MsgBox "Name not found, check spelling", , "Error"
    End If
End Sub

Private Sub cmdjersey_Click()
'view roster by jersey# numerically
ctr = 32
picresults.Cls

For pass = 1 To ctr - 1
    For comp = 1 To ctr - 1
        If jersey(comp) > jersey(comp + 1) Then
        
        'switch assists
        tempassists = assists(comp)
        assists(comp) = assists(comp + 1)
        assists(comp + 1) = tempassists
        
        'switch goals
        tempgoals = goals(comp)
        goals(comp) = goals(comp + 1)
        goals(comp + 1) = tempgoals
        
        'switch names
        tempnames = names(comp)
        names(comp) = names(comp + 1)
        names(comp + 1) = tempnames
        
        'switch jersey
        tempjersey = jersey(comp)
        jersey(comp) = jersey(comp + 1)
        jersey(comp + 1) = tempjersey
        
        'switch games played
        tempgames = games(comp)
        games(comp) = games(comp + 1)
        games(comp + 1) = tempgames
        
        
        'switch points
        temppoints = points(comp)
        points(comp) = points(comp + 1)
        points(comp + 1) = temppoints
        
        'switch penalty
        temppenalty = penalty(comp)
        penalty(comp) = penalty(comp + 1)
        penalty(comp + 1) = temppenalty
        
        'switch plus minus
        tempplusminus = plusminus(comp)
        plusminus(comp) = plusminus(comp + 1)
        plusminus(comp + 1) = tempplusminus
        
        End If
        Next comp
        Next pass
        
     
picresults.Print ; "Player                            jersey#   Games-played   Goals       Assists       Points    Penalty Minutes      +/- "
picresults.Print
For x = 1 To 32
picresults.Print names(x);
picresults.Print Tab(25); jersey(x);
picresults.Print Tab(37); games(x); Tab(50); goals(x); Tab(60); assists(x); Tab(70); points(x); Tab(80); penalty(x); Tab(95); plusminus(x)
   
        Next x
End Sub

Private Sub cmdnextform_Click()
Form1.Hide  'hides original and shows second form
Form2.Show
End Sub

Private Sub cmdpenalty_Click()
'view roster by total penalty minutes
ctr = 32
picresults.Cls

For pass = 1 To ctr - 1
    For comp = 1 To ctr - 1
        If penalty(comp) < penalty(comp + 1) Then
        
        'switch assists
        tempassists = assists(comp)
        assists(comp) = assists(comp + 1)
        assists(comp + 1) = tempassists
        
        'switch goals
        tempgoals = goals(comp)
        goals(comp) = goals(comp + 1)
        goals(comp + 1) = tempgoals
        
        'switch names
        tempnames = names(comp)
        names(comp) = names(comp + 1)
        names(comp + 1) = tempnames
        
        'switch jersey
        tempjersey = jersey(comp)
        jersey(comp) = jersey(comp + 1)
        jersey(comp + 1) = tempjersey
        
        'switch games played
        tempgames = games(comp)
        games(comp) = games(comp + 1)
        games(comp + 1) = tempgames
        
        
        'switch points
        temppoints = points(comp)
        points(comp) = points(comp + 1)
        points(comp + 1) = temppoints
        
        'switch penalty
        temppenalty = penalty(comp)
        penalty(comp) = penalty(comp + 1)
        penalty(comp + 1) = temppenalty
        
        'switch plus minus
        tempplusminus = plusminus(comp)
        plusminus(comp) = plusminus(comp + 1)
        plusminus(comp + 1) = tempplusminus
        
        End If
        Next comp
        Next pass
        
     
picresults.Print ; "Player                            jersey#   Games-played   Goals       Assists       Points    Penalty Minutes      +/- "
picresults.Print
For x = 1 To 32
picresults.Print names(x);
picresults.Print Tab(25); jersey(x);
picresults.Print Tab(37); games(x); Tab(50); goals(x); Tab(60); assists(x); Tab(70); points(x); Tab(80); penalty(x); Tab(95); plusminus(x)
            
        Next x
End Sub

Private Sub cmdplusminus_Click()
'view arranged by plus minus
ctr = 32
picresults.Cls

For pass = 1 To ctr - 1
    For comp = 1 To ctr - 1
        If plusminus(comp) < plusminus(comp + 1) Then
        
        'switch assists
        tempassists = assists(comp)
        assists(comp) = assists(comp + 1)
        assists(comp + 1) = tempassists
        
        'switch goals
        tempgoals = goals(comp)
        goals(comp) = goals(comp + 1)
        goals(comp + 1) = tempgoals
        
        'switch names
        tempnames = names(comp)
        names(comp) = names(comp + 1)
        names(comp + 1) = tempnames
        
        'switch jersey
        tempjersey = jersey(comp)
        jersey(comp) = jersey(comp + 1)
        jersey(comp + 1) = tempjersey
        
        'switch games played
        tempgames = games(comp)
        games(comp) = games(comp + 1)
        games(comp + 1) = tempgames
        
        
        'switch points
        temppoints = points(comp)
        points(comp) = points(comp + 1)
        points(comp + 1) = temppoints
        
        'switch penalty
        temppenalty = penalty(comp)
        penalty(comp) = penalty(comp + 1)
        penalty(comp + 1) = temppenalty
        
        'switch plus minus
        tempplusminus = plusminus(comp)
        plusminus(comp) = plusminus(comp + 1)
        plusminus(comp + 1) = tempplusminus
        
        End If
        Next comp
        Next pass
        
     
picresults.Print ; "Player                            jersey#   Games-played   Goals       Assists       Points    Penalty Minutes      +/- "
picresults.Print
For x = 1 To 32
picresults.Print names(x);
picresults.Print Tab(25); jersey(x);
picresults.Print Tab(37); games(x); Tab(50); goals(x); Tab(60); assists(x); Tab(70); points(x); Tab(80); penalty(x); Tab(95); plusminus(x)
            
        Next x
        
End Sub

Private Sub cmdpoints_Click()
'view roster arranged by points
ctr = 32
picresults.Cls

For pass = 1 To ctr - 1
    For comp = 1 To ctr - 1
        If points(comp) < points(comp + 1) Then
        
        'switch assists
        tempassists = assists(comp)
        assists(comp) = assists(comp + 1)
        assists(comp + 1) = tempassists
        
        'switch goals
        tempgoals = goals(comp)
        goals(comp) = goals(comp + 1)
        goals(comp + 1) = tempgoals
        
        'switch names
        tempnames = names(comp)
        names(comp) = names(comp + 1)
        names(comp + 1) = tempnames
        
        'switch jersey
        tempjersey = jersey(comp)
        jersey(comp) = jersey(comp + 1)
        jersey(comp + 1) = tempjersey
        
        'switch games played
        tempgames = games(comp)
        games(comp) = games(comp + 1)
        games(comp + 1) = tempgames
        
        
        'switch points
        temppoints = points(comp)
        points(comp) = points(comp + 1)
        points(comp + 1) = temppoints
        
        'switch penalty
        temppenalty = penalty(comp)
        penalty(comp) = penalty(comp + 1)
        penalty(comp + 1) = temppenalty
        
        'switch plus minus
        tempplusminus = plusminus(comp)
        plusminus(comp) = plusminus(comp + 1)
        plusminus(comp + 1) = tempplusminus
        
        End If
        Next comp
        Next pass
        
     
picresults.Print ; "Player                            jersey#   Games-played   Goals       Assists       Points    Penalty Minutes      +/- "
picresults.Print
For x = 1 To 32
picresults.Print names(x);
picresults.Print Tab(25); jersey(x);
picresults.Print Tab(37); games(x); Tab(50); goals(x); Tab(60); assists(x); Tab(70); points(x); Tab(80); penalty(x); Tab(95); plusminus(x)
            
        Next x
        
End Sub

Private Sub cmdquit_Click()
End             'ends program
End Sub

Private Sub cmdread_Click()
'open file and read stats into an array
'Open "M:\CS130\VB project\wildstatsnew.txt" For Input As #1
Open path & "wildstatsnew.txt" For Input As #1

picresults.Cls
For x = 1 To 32
    Input #1, names(x), games(x), goals(x), assists(x), points(x), penalty(x), plusminus(x), jersey(x)
Next x

'print stats and names
picresults.Print "Player                            jersey#   Games-played   Goals       Assists       Points    Penalty Minutes      +/- "
picresults.Print
For x = 1 To 32
picresults.Print names(x);
picresults.Print Tab(25); jersey(x);
picresults.Print Tab(37); games(x); Tab(50); goals(x); Tab(60); assists(x); Tab(70); points(x); Tab(80); penalty(x); Tab(95); plusminus(x)
Next x
cmdjersey.Enabled = True
cmdgames.Enabled = True
cmdgoals.Enabled = True
cmdassists.Enabled = True
cmdpoints.Enabled = True
cmdpenalty.Enabled = True
cmdplusminus.Enabled = True
cmdindividual.Enabled = True
cmdnextform.Enabled = True


Close (1)
End Sub

Private Sub Form_Load()
path = "M:\CS130\VB project\"
LoadPicture (path & "jerseysnew.jpg")
End Sub
