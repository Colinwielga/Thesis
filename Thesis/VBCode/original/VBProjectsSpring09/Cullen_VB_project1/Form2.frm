VERSION 5.00
Begin VB.Form frmstats 
   BackColor       =   &H000000C0&
   Caption         =   "Form2"
   ClientHeight    =   9615
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   ScaleHeight     =   9615
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdper48 
      Caption         =   "Load Stats per 48 minutes"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.PictureBox picresults2 
      BackColor       =   &H8000000E&
      Height          =   4335
      Left            =   6360
      ScaleHeight     =   4275
      ScaleWidth      =   6075
      TabIndex        =   9
      Top             =   1920
      Width           =   6135
   End
   Begin VB.CommandButton cmdminutes 
      Caption         =   "Sort by minutes"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "quit"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      TabIndex        =   7
      Top             =   5760
      Width           =   3135
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   1215
      Left            =   6360
      ScaleHeight     =   1155
      ScaleWidth      =   6075
      TabIndex        =   6
      Top             =   480
      Width           =   6135
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to main page"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdassists 
      Caption         =   "Sort By Assists"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdrebounds 
      Caption         =   "Sort by Rebounds"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdpoints 
      Caption         =   "Sort by Points Per Game"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdplayer 
      Caption         =   "Search for a Specific Player"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load Stats Per Game"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmstats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chicago Bulls (Chicagobulls.vbp)
    'frmstatsfrmBullsplayers.frm)
    'Written by: Brian Cullen
    'Written on: March 16, 2008
    'Objective: This form allows the user to sort the Chicago Bulls players by a certain
    'statistical category as well as searching for a specific player's statistics.
    'viewing a photo of each player.


Option Explicit
Dim player(1 To 14) As String
Dim PPG(1 To 14) As Single
Dim rebounds(1 To 14) As Single
Dim assists(1 To 14) As Single
Dim Minutes(1 To 14) As Single
Dim ctr As Integer
Dim Temp As Single, I As Integer
Dim rofl As String, temptwo As String, tempthree As Single, tempfour As Single, tempfive As Single





Private Sub cmdload_Click()
picresults.Cls
picresults2.Cls
'This opens the file
Open App.Path & "\playerstats.txt" For Input As #1
'this reads the files into five parallel arrays
picresults2.Print "Name             PPG            rebounds              assists            minutes"
ctr = 0
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, player(ctr), PPG(ctr), rebounds(ctr), assists(ctr), Minutes(ctr)
    picresults2.Print player(ctr), PPG(ctr), rebounds(ctr), assists(ctr), Minutes(ctr)
    
Loop
picresults.Print "Player listed alphabetically with statistics per game"
Close #1

End Sub


Private Sub cmdminutes_Click()
Dim pass As Integer, pos As Integer
Dim I As Integer, temp2 As String, tempthree As Single, tempfour As Single, tempfive As Single
picresults.Cls
picresults2.Cls
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If Minutes(pos) < Minutes(pos + 1) Then
            Temp = Minutes(pos)
            Minutes(pos) = Minutes(pos + 1)
            Minutes(pos + 1) = Temp
            temp2 = player(pos)
            player(pos) = player(pos + 1)
            player(pos + 1) = temp2
            tempthree = PPG(pos)
            PPG(pos) = PPG(pos + 1)
            PPG(pos + 1) = tempthree
            tempfour = rebounds(pos)
            rebounds(pos) = rebounds(pos + 1)
            rebounds(pos + 1) = tempfour
            tempfive = assists(pos)
            assists(pos) = assists(pos + 1)
            assists(pos) = tempfive
        End If
    Next pos
Next pass
 picresults.Print "Players sorted by minutes Per Game"
 picresults2.Print "NAME"; Tab(20); "MPG"
For I = 1 To ctr
    picresults2.Print player(I); Tab(20); Minutes(I)
    Next I
    
End Sub



Private Sub cmdper48_Click()
picresults.Cls
picresults2.Cls
'This opens the file
Open App.Path & "\playerstats.txt" For Input As #1
'this reads the files into five parallel arrays
picresults2.Print "Name             PPG            rebounds              assists"
ctr = 0
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, player(ctr), PPG(ctr), rebounds(ctr), assists(ctr), Minutes(ctr)
    picresults2.Print player(ctr), FormatNumber(48 / Minutes(ctr) * PPG(ctr), 1), FormatNumber(48 / Minutes(ctr) * rebounds(ctr), 1), FormatNumber(48 / Minutes(ctr) * assists(ctr), 1)
    
Loop
picresults.Print "Player listed alphabetically with statistics "
picresults.Print "per 48 minutes"
Close #1

End Sub

Private Sub cmdplayer_Click()
Dim Found As Boolean
Dim I As Integer
Dim searchplayer As String
picresults.Cls
picresults2.Cls
searchplayer = InputBox("Enter the name of the player you wish to find", "Name")
I = 0
Found = False

Do While ((Not Found) And (I < ctr))
    I = I + 1
    If searchplayer = player(I) Then
        Found = True
    End If
Loop

If (Not Found) Then
    picresults.Print searchplayer; "was not on the '95-'96 Chicago Bulls"
Else
picresults.Print "points rebounds, assists, and minutes per game of a specific player"
picresults2.Print "Name"; Tab(8); "PPG"; Tab(16); "Rebs"; Tab(24); "Assists"; Tab(32); "MPG"
picresults2.Print searchplayer; Tab(8); PPG(I); Tab(16); rebounds(I); Tab(24); assists(I); Tab(32); Minutes(I)
End If

End Sub

Private Sub cmdpoints_Click()

Dim pass As Integer, pos As Integer
Dim I As Integer, temp2 As String, tempthree As Single, tempfour As Single, tempfive As Single
picresults.Cls
picresults2.Cls

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If PPG(pos) < PPG(pos + 1) Then
            Temp = PPG(pos)
            PPG(pos) = PPG(pos + 1)
            PPG(pos + 1) = Temp
            temp2 = player(pos)
            player(pos) = player(pos + 1)
            player(pos + 1) = temp2
            tempthree = Minutes(pos)
            Minutes(pos) = Minutes(pos + 1)
            Minutes(pos + 1) = tempthree
            tempfour = rebounds(pos)
            rebounds(pos) = rebounds(pos + 1)
            rebounds(pos + 1) = tempfour
            tempfive = assists(pos)
            assists(pos) = assists(pos + 1)
            assists(pos) = tempfive
        End If
    Next pos
Next pass
    picresults.Print "Players sorted by Points Per Game"
    picresults2.Print "Name                    PPG"
For I = 1 To ctr
    picresults2.Print player(I); Tab(20); PPG(I)
    Next I
    
    
End Sub

Private Sub cmdrebounds_Click()

Dim pass As Integer, pos As Integer
Dim I As Integer, temp2 As String, tempthree As Single, tempfour As Single, tempfive As Single
picresults.Cls
picresults2.Cls
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If rebounds(pos) < rebounds(pos + 1) Then
            Temp = rebounds(pos)
            rebounds(pos) = rebounds(pos + 1)
            rebounds(pos + 1) = Temp
            temp2 = player(pos)
            player(pos) = player(pos + 1)
            player(pos + 1) = temp2
            tempthree = PPG(pos)
            PPG(pos) = PPG(pos + 1)
            PPG(pos + 1) = tempthree
            tempfour = Minutes(pos)
            Minutes(pos) = Minutes(pos + 1)
            Minutes(pos + 1) = tempfour
            tempfive = assists(pos)
            assists(pos) = assists(pos + 1)
            assists(pos) = tempfive
        End If
    Next pos
Next pass
picresults2.Print "Name"; Tab(20); "RPG"
For I = 1 To ctr
    picresults2.Print player(I); Tab(20); FormatNumber(rebounds(I), 1)
    Next I
    
   picresults.Print "Players sorted by rebounds per game"
End Sub

Private Sub cmdassists_Click()

Dim pass As Integer, pos As Integer
Dim I As Integer, temp2 As String, tempthree As Single, tempfour As Single, tempfive As Single
picresults.Cls
picresults2.Cls
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If assists(pos) < assists(pos + 1) Then
            Temp = assists(pos)
            assists(pos) = assists(pos + 1)
            assists(pos + 1) = Temp
            temp2 = player(pos)
            player(pos) = player(pos + 1)
            player(pos + 1) = temp2
            tempthree = PPG(pos)
            PPG(pos) = PPG(pos + 1)
            PPG(pos + 1) = tempthree
            tempfour = rebounds(pos)
            rebounds(pos) = rebounds(pos + 1)
            rebounds(pos + 1) = tempfour
            tempfive = Minutes(pos)
            Minutes(pos) = Minutes(pos + 1)
            Minutes(pos) = tempfive
        End If
    Next pos
Next pass

picresults.Print "Players sorted by assists per game"
picresults2.Print "Name"; Tab(20); "APG"
For I = 1 To ctr
    picresults2.Print player(I); Tab(20); assists(I)
    Next I
    
End Sub

Private Sub cmdreturn_Click()
frmstats.Hide
frmmainpage.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub


