VERSION 5.00
Begin VB.Form frmStats 
   Caption         =   "Stats"
   ClientHeight    =   7380
   ClientLeft      =   2715
   ClientTop       =   1830
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   9975
   Begin VB.CommandButton cmdSort3 
      BackColor       =   &H0000C000&
      Caption         =   "Sort Stats By Money"
      Height          =   615
      Left            =   4200
      MaskColor       =   &H00008080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdSort2 
      BackColor       =   &H000080FF&
      Caption         =   "Sort Stats By Name"
      Height          =   615
      Left            =   2160
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back to Lobby"
      Height          =   615
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H000040C0&
      Caption         =   "Show Stats"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000013&
      Height          =   6015
      Left            =   480
      ScaleHeight     =   5955
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   10710
      Left            =   0
      Picture         =   "frmStats.frx":0000
      Top             =   0
      Width           =   15000
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Mystake Lake Casino
'Authors: David Johnson And Jeremy Iverson
'Date: Monday, November 5, 2007

Option Explicit
'This form allows the user to view how much other players have won or lost
Dim names(1 To 100)
Dim money(1 To 100)
Dim ctr As Integer
Dim click As Integer

Private Sub cmdQuit_Click()
    'Go back to Lobby
    frmStats.Hide
    frmLobby.Show
End Sub

Private Sub cmdShow_Click()
    'Reads data from a file and displays the data in picturebox
    picResults.Cls
    picResults.Print "Players", Tab(30), "Winnings"
    picResults.Print "*****************************************************************"
    ctr = 0
    Open App.Path & "\standings.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, names(ctr), money(ctr)
        picResults.Print names(ctr), Tab(30), FormatCurrency(money(ctr))
    Loop
    Close #1
End Sub

Private Sub cmdSort2_Click()
    'Sorts the data by name in asending and desending order,
    'depending on the number of clicks by the user
    Dim pos As Integer
    Dim pass As Integer
    Dim tempn As String
    Dim tempm As Single
    picResults.Cls
    picResults.Print "Players", Tab(30), "Winnings"
    picResults.Print "*****************************************************************"
    click = click + 1
    If (click / 2) = Int(click / 2) Then
        For pass = 1 To ctr - 1
            For pos = 1 To ctr - pass
                If names(pos) > names(pos + 1) Then
                    tempn = names(pos)
                    names(pos) = names(pos + 1)
                    names(pos + 1) = tempn
                    tempm = money(pos)
                    money(pos) = money(pos + 1)
                    money(pos + 1) = tempm
                End If
            Next pos
        Next pass
    Else
        For pass = 1 To ctr - 1
            For pos = 1 To ctr - pass
                If names(pos) < names(pos + 1) Then
                    tempn = names(pos)
                    names(pos) = names(pos + 1)
                    names(pos + 1) = tempn
                    tempm = money(pos)
                    money(pos) = money(pos + 1)
                    money(pos + 1) = tempm
                End If
            Next pos
        Next pass
    End If
    For pos = 1 To ctr
        picResults.Print names(pos), Tab(30), FormatCurrency(money(pos))
    Next pos
End Sub

Private Sub cmdSort3_Click()
    'Sorts the data by amount of money in asending and desending order,
    'depending on the number of clicks by the user
    Dim pos As Integer
    Dim pass As Integer
    Dim tempn As String
    Dim tempm As Single
    picResults.Cls
    picResults.Print "Players", Tab(30), "Winnings"
    picResults.Print "*****************************************************************"
    click = click + 1
    If (click / 2) = Int(click / 2) Then
        For pass = 1 To ctr - 1
            For pos = 1 To ctr - pass
                If money(pos) < money(pos + 1) Then
                    tempn = names(pos)
                    names(pos) = names(pos + 1)
                    names(pos + 1) = tempn
                    tempm = money(pos)
                    money(pos) = money(pos + 1)
                    money(pos + 1) = tempm
                End If
            Next pos
        Next pass
    Else
        For pass = 1 To ctr - 1
            For pos = 1 To ctr - pass
                If money(pos) > money(pos + 1) Then
                    tempn = names(pos)
                    names(pos) = names(pos + 1)
                    names(pos + 1) = tempn
                    tempm = money(pos)
                    money(pos) = money(pos + 1)
                    money(pos + 1) = tempm
                End If
            Next pos
        Next pass
    End If
    For pos = 1 To ctr
        picResults.Print names(pos), Tab(30), FormatCurrency(money(pos))
    Next pos
End Sub

