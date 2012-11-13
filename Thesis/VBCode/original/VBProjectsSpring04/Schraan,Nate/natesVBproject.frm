VERSION 5.00
Begin VB.Form TeamBlueProject 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   3735
   ClientTop       =   1140
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   9915
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   120
      Picture         =   "NATESV~1.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   4515
      TabIndex        =   11
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton cmdfindperson 
      BackColor       =   &H00404000&
      Caption         =   "Find a Player"
      Height          =   975
      Left            =   7320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdslugaverage 
      BackColor       =   &H000080FF&
      Caption         =   "Slugging Percentage"
      Height          =   975
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdindba 
      BackColor       =   &H00C0C000&
      Caption         =   "Find Players Individual Batting Average"
      Height          =   975
      Left            =   7320
      MaskColor       =   &H00808000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdinputbox 
      BackColor       =   &H00800080&
      Caption         =   "Players with a Batting Average Above"
      Height          =   975
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdsorttriples 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort Team by Number of Triples"
      Height          =   975
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H0080C0FF&
      Caption         =   " Sort Team by Home Runs"
      Height          =   975
      Left            =   4920
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdbataverage 
      BackColor       =   &H0000FFFF&
      Caption         =   "Calculate Blue Teams Batting Average"
      Height          =   975
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton Quit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton teambluedata 
      BackColor       =   &H00FF8080&
      Caption         =   "Load Blue Teams Data into program"
      Height          =   735
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   240
      ScaleHeight     =   4995
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "First press the Blue Button"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "TeamBlueProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim player(1 To 100) As String              'Nate Schraan
Dim ave As Single                           'schraansVbproject(M:\cs130\VB Project\SchraansVbproject)
Dim hits(1 To 100) As Single                'natesVbproject(M:\cs130\VB Project\natesVBproject.frm)
Dim atbats(1 To 100) As Single              'Author Nate Schraan
Dim hr(1 To 100) As Integer                 'Date Written March 10, 2004
Dim s(1 To 100) As Integer                  'purpose to find statistics such as batting average slugging percentage and to organize your team by certain statistics
Dim doubles(1 To 100) As Integer
Dim triples(1 To 100) As Integer    'option explicit uses global variables throughout the program
Dim PATH As String
Dim ctr As Integer

Private Sub cmdbataverage_Click()   'calculates the total team batting average
picresults.Cls
Dim h As Integer
Dim ab As Integer
Dim t As Single
For J = 1 To 9
h = h + hits(J)
ab = ab + atbats(J)
Next J
t = h / ab
picresults.Print "----BLUE TEAMS BATTING AVERAGE----"
picresults.Print FormatNumber(t, 3)
End Sub
Private Sub cmdfind_Click()
picresults.Cls
Dim tempplayer As String
Dim temphits As Integer
Dim tempatbats As Integer
Dim temphr As Integer
Dim temps As Integer
Dim tempdoubles As Integer
Dim temptriples As Integer
Dim pass As Integer
Dim comp As Integer
Dim J As Integer
For pass = 1 To ctr
    For comp = 1 To ctr - pass
    If hr(comp) < hr(comp + 1) Then
        tempplayer = player(comp)       'sorts names
        player(comp) = player(comp + 1)
        player(comp + 1) = tempplayer
        temphits = hits(comp)           'sorts hits
        hits(comp) = hits(comp + 1)
        hits(comp + 1) = temphits
        tempatbats = atbats(comp)       'sorts at bats
        atbats(comp) = atbats(comp + 1)
        atbats(comp + 1) = tempatbats
        temphr = hr(comp)               'sorts homeruns
        hr(comp) = hr(comp + 1)
        hr(comp + 1) = temphr
        temps = s(comp)                 'sorts singles
        s(comp) = s(comp + 1)
        s(comp + 1) = temps
        tempdoubles = doubles(comp)     'sorts doubles
        doubles(comp) = doubles(comp + 1)
        doubles(comp + 1) = tempdoubles
        temptriples = triples(comp)
        triples(comp) = triples(comp + 1)
        triples(comp + 1) = temptriples
        
    End If
    Next comp
Next pass
picresults.Print "TEAM SORTED BY NUMBER OF HOMERUNS"
For J = 1 To ctr
picresults.Print player(J); Tab(18); hr(J)
Next J
End Sub

Private Sub cmdfindperson_Click() 'finds a person that is given by the user in an input box
picresults.Cls
Dim found As Boolean
Dim I As Integer
Dim n As String
Dim t As String
n = InputBox("Enter First and Last Name of Player You Wish to Find", "Name")
I = 0
found = False
Do While (notfound) Or I <= 10
    I = I + 1
    If n = player(I) Then
    picresults.Print "Player Found"
    picresults.Print n
    found = True
    End If
Loop
If found = False Then picresults.Print "Player Not Found"
End Sub

Private Sub cmdindba_Click()    'calculates each batting average for each player
picresults.Cls
Dim y As Single
picresults.Print "INDIVIDUAL BATTING AVERAGES"
For J = 1 To ctr
    y = hits(J) / atbats(J)
    picresults.Print player(J); Tab(18); FormatNumber(y, 3)
Next J


End Sub

Private Sub cmdinputbox_Click() 'takes data from user from an input box and lists all players with batting averages above that number
picresults.Cls
Dim n As Single
Dim y As Single
n = InputBox("Enter Batting Average .100 to.500")
 picresults.Print "PLAYERS WITH BATTING AVERAGES OVER"; Tab(47); FormatNumber(n, 3)
 For J = 1 To ctr
    If hits(J) / atbats(J) > n Then
    y = hits(J) / atbats(J)
    picresults.Print player(J); Tab(18); FormatNumber(y, 3)
    End If
Next J
 If n > y Then
    MsgBox "Sorry There are No Players with an Average Above Your Number"
    End If
End Sub

Private Sub cmdslugaverage_Click()   'calculates each teammembers slugging percentage
picresults.Cls
picresults.Print "-----SLUGGING PERCENTAGE------"
Dim y As Single
For J = 1 To ctr
y = s(J) * 1 + doubles(J) * 2 + triples(J) * 3 + hr(J) * 4
t = y / atbats(J)
picresults.Print player(J); Tab(18); FormatNumber(t, 3)
Next J
End Sub

Private Sub cmdsorttriples_Click() 'sorts through array and displays the number of triples each player has
picresults.Cls
For pass = 1 To ctr
    For comp = 1 To ctr - pass
    If triples(comp) < triples(comp + 1) Then
        tempplayer = player(comp)       'sorts names
        player(comp) = player(comp + 1)
        player(comp + 1) = tempplayer
        temphits = hits(comp)           'sorts hits
        hits(comp) = hits(comp + 1)
        hits(comp + 1) = temphits
        tempatbats = atbats(comp)       'sorts at bats
        atbats(comp) = atbats(comp + 1)
        atbats(comp + 1) = tempatbats
        temphr = hr(comp)               'sorts homeruns
        hr(comp) = hr(comp + 1)
        hr(comp + 1) = temphr
        temps = s(comp)                 'sorts singles
        s(comp) = s(comp + 1)
        s(comp + 1) = temps
        tempdoubles = doubles(comp)     'sorts doubles
        doubles(comp) = doubles(comp + 1)
        doubles(comp + 1) = tempdoubles
        temptriples = triples(comp)     'sorts triples
        triples(comp) = triples(comp + 1)
        triples(comp + 1) = temptriples
        
    End If
    Next comp
Next pass
picresults.Print "PLAYERS IN ORDER OF TRIPLES"
For J = 1 To ctr
picresults.Print player(J); Tab(18); triples(J)
Next J
End Sub

Private Sub Form_Load()
PATH = "M:\cs130\Labs\Lab10\"
End Sub

Private Sub Quit_Click()
End
End Sub

Private Sub teambluedata_Click() 'loads array file into program
ctr = 0
Open PATH & "teamblue.txt" For Input As #1
Do While Not EOF(1)
ctr = ctr + 1
Input #1, player(ctr), hits(ctr), atbats(ctr), hr(ctr), s(ctr), doubles(ctr), triples(ctr)
Loop
teambluedata.Enabled = False
Close #1
End Sub

