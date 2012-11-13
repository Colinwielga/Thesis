VERSION 5.00
Begin VB.Form Lacrosse 
   BackColor       =   &H00000080&
   Caption         =   "Lacrosse"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Search 
      Caption         =   "Search by player"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton AverageGoals 
      Caption         =   "Average Goals"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Sort 
      Caption         =   "Sort  by Name"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load and Print"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   1455
   End
   Begin VB.PictureBox pbxResults 
      Height          =   2895
      Left            =   2280
      ScaleHeight     =   2835
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   2760
      Width           =   7335
   End
   Begin VB.Image LaxStick 
      Height          =   3000
      Left            =   4320
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "Lacrosse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'LacrosseProject (Project1)
'Lacrosse
'Anne Mills
'10/22/03
'The purpose is to easily access information about the top players in CSB Women's Lacrosse

Option Explicit
Dim i As Integer
Dim strName(1 To 5) As String
Dim game1(1 To 4), game2(1 To 4), game3(1 To 4), game4(1 To 4), game5(1 To 4) As Integer
Dim temp As String
Dim temp1, temp2, temp3, temp4, temp5 As Integer
Dim strPath As String
Dim strFile As String


Private Sub AverageGoals_Click() 'Print the Average goals of each player from the 5 games
pbxResults.Print
Dim AverageGoals As Single
For i = 1 To 4
    AverageGoals = (game1(i) + game2(i) + game3(i) + game4(i) + game5(i)) / 5
pbxResults.Print strName(i); " has an average goals per game of "; AverageGoals
Next i
End Sub


Private Sub Load_Click() 'This is the starter button, it loads and prints the begining information, unaltered.

strPath = "N:\CS130\handin\Project 1\" 'This can be changed to any address
strFile = strPath & "players' stats.txt"
Open strFile For Input As #1
    For i = 1 To 4
        Input #1, strName(i), game1(i), game2(i), game3(i), game4(i), game5(i)
    Next i
Close #1

pbxResults.Cls
pbxResults.Print "Goals scored by the top four women's lacrosse players in 2003"
pbxResults.Print
pbxResults.Print
pbxResults.Print "Player", "Game 1", "Game 2", "Game 3", "Game 4", "Game 5"
pbxResults.Print
For i = 1 To 4
    pbxResults.Print strName(i), game1(i), game2(i), game3(i), game4(i), game5(i)
Next i
End Sub


Private Sub Quit_Click()
End
End Sub

Private Sub Search_Click() 'Search and print for one player's stats at a time
Dim Player As String
Dim Found As Boolean
Dim AverageG As Single

Found = False
i = 0
Player = InputBox("Enter Name")
pbxResults.Cls
Do While i < 5 And Found = False
    i = i + 1 'counter
    If Player = strName(i) Then
        Found = True
    End If
Loop

If Found = True Then
   
AverageG = (game1(i) + game2(i) + game3(i) + game4(i) + game5(i)) / 5
pbxResults.Print "Player", "Game 1", "Game 2", "Game 3", "Game 4", "Game 5"
   
    pbxResults.Print
    pbxResults.Print strName(i), game1(i), game2(i), game3(i), game4(i), game5(i)
    pbxResults.Print
    pbxResults.Print strName(i), "has an average of "; AverageG; " goals per game"
Else
    pbxResults.Print "Name not found"
End If
End Sub

Private Sub Sort_Click() 'sorts the players and their stats by their name

Dim pass As Integer
Dim n As Integer
n = 4

For pass = 1 To n - 1
    For i = 1 To n - pass
        If strName(i) > strName(i + 1) Then 'sorts names
            temp = strName(i + 1)
            strName(i + 1) = strName(i)
            strName(i) = temp
            temp1 = game1(i + 1) 'sorts stats
            game1(i + 1) = game1(i)
            game1(i) = temp1
            temp2 = game2(i + 1)
            game2(i + 1) = game2(i)
            game2(i) = temp2
            temp3 = game3(i + 1)
            game3(i + 1) = game3(i)
            game3(i) = temp3
            temp4 = game4(i + 1)
            game4(i + 1) = game4(i)
            game4(i) = temp4
            temp5 = game5(i + 1)
            game5(i + 1) = game5(i)
            game5(i) = temp5
        End If
    Next i
Next pass

pbxResults.Cls
pbxResults.Print "Goals scored by the top four women's lacrosse players in 2003"
pbxResults.Print
pbxResults.Print
pbxResults.Print "Player", "Game 1", "Game 2", "Game 3", "Game 4", "Game 5"
pbxResults.Print
For i = 1 To 4
    pbxResults.Print strName(i), game1(i), game2(i), game3(i), game4(i), game5(i)
Next i
End Sub
