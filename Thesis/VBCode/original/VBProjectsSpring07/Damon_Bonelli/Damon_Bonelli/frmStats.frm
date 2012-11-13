VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H000000FF&
   Caption         =   "Statistics"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      Height          =   4575
      Left            =   480
      ScaleHeight     =   4515
      ScaleWidth      =   8355
      TabIndex        =   7
      Top             =   3240
      Width           =   8415
   End
   Begin VB.CommandButton cmdkillpercentage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compute Kill Percentage"
      Height          =   855
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdBlock 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find Blocking Leaders"
      Height          =   855
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdService 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Service Statistics"
      Height          =   855
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdkpleaders 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find Kill Percentage Leaders"
      Height          =   855
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find a Player"
      Height          =   855
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdhome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Home"
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBlock_Click()
'This subroutine determines the blocking leaders on the team, and then
'asks the user how many players he/she would like to see (i.e. if
'the user were to input 3, the 3 blockers with the most blocks
'would be displayed in the picture box.  This button uses a bubble sort
'method.  If the user chooses a number of players greater than the number
'of players on the team (in the file), it will return an error message.
picresults.Cls
Dim pass As Single, pos As Single, n As Single, temp As Single, number As Single
Dim tempn As String
n = InputBox("How many players would you like to see?", "Blockers")
    For pass = 1 To ctr - 1
        For pos = 1 To ctr - pass
            If blocks(pos) > blocks(pos + 1) Then
                temp = blocks(pos + 1)
                blocks(pos + 1) = blocks(pos)
                blocks(pos) = temp
                tempn = names(pos + 1)
                names(pos + 1) = names(pos)
                names(pos) = tempn
                temp = attempts(pos + 1)
                attempts(pos + 1) = attempts(pos)
                attempts(pos) = temp
                temp = kills(pos + 1)
                kills(pos + 1) = kills(pos)
                kills(pos) = temp
                temp = errors(pos + 1)
                errors(pos + 1) = errors(pos)
                errors(pos) = temp
                temp = aces(pos + 1)
                aces(pos + 1) = aces(pos)
                aces(pos) = temp
                temp = se(pos + 1)
                se(pos + 1) = se(pos)
                se(pos) = temp
                temp = games(pos + 1)
                games(pos + 1) = games(pos)
                games(pos) = temp
                temp = jersey(pos + 1)
                jersey(pos + 1) = jersey(pos)
                jersey(pos) = temp
                tempn = position(pos + 1)
                position(pos + 1) = position(pos)
                position(pos) = tempn
            End If
        Next pos
    Next pass
If n <= ctr Then
    Do Until number >= n
        picresults.Print names(ctr - number), " : "; blocks(ctr - number); " blocks."
        number = number + 1
    Loop
End If
If n > ctr Then
    MsgBox "You selected too many players.", , "Error"
End If
End Sub

Private Sub cmdend_Click()
'quits the program
End
End Sub



Private Sub cmdhome_Click()
'goes to the statistics screen
frmintro.Show
frmStats.Hide
End Sub

Private Sub cmdkpleaders_Click()
'This subroutine determines the kill percentage leaders on the team
'by first calculating the kill percentage for every player, then
'asking the user how many players he/she would like to see (i.e. if
'the user were to input 3, the 3 hitters with the highest kill percentage
'would be displayed in the picture box.  This button uses a bubble sort
'method.  If the user chooses a number of players greater than the number
'of players on the team (in the file), it will return an error message.
Dim pass As Single, pos As Single, n As Single, temp As Single, number As Single
Dim tempn As String, hold As Single
picresults.Cls
Dim killpercentage(1 To 30) As Single
Do Until hold >= ctr
    hold = hold + 1
    killpercentage(hold) = (kills(hold) - errors(hold)) / attempts(hold)
Loop
n = InputBox("How many players would you like to see?", "Hitters")
    For pass = 1 To ctr - 1
        For pos = 1 To ctr - pass
            If killpercentage(pos) > killpercentage(pos + 1) Then
                temp = killpercentage(pos + 1)
                killpercentage(pos + 1) = killpercentage(pos)
                killpercentage(pos) = temp
                temp = blocks(pos + 1)
                blocks(pos + 1) = blocks(pos)
                blocks(pos) = temp
                tempn = names(pos + 1)
                names(pos + 1) = names(pos)
                names(pos) = tempn
                temp = attempts(pos + 1)
                attempts(pos + 1) = attempts(pos)
                attempts(pos) = temp
                temp = kills(pos + 1)
                kills(pos + 1) = kills(pos)
                kills(pos) = temp
                temp = errors(pos + 1)
                errors(pos + 1) = errors(pos)
                errors(pos) = temp
                temp = aces(pos + 1)
                aces(pos + 1) = aces(pos)
                aces(pos) = temp
                temp = se(pos + 1)
                se(pos + 1) = se(pos)
                se(pos) = temp
                temp = games(pos + 1)
                games(pos + 1) = games(pos)
                games(pos) = temp
                temp = jersey(pos + 1)
                jersey(pos + 1) = jersey(pos)
                jersey(pos) = temp
                tempn = position(pos + 1)
                position(pos + 1) = position(pos)
                position(pos) = tempn
            End If
        Next pos
    Next pass
If n <= ctr Then
    Do Until number >= n
        picresults.Print names(ctr - number), " : "; FormatNumber(killpercentage(ctr - number), 3)
        number = number + 1
    Loop
End If
If n > ctr Then
    MsgBox "You selected too many players.", , "Error"
End If
End Sub

Private Sub cmdQuit_Click()
'quits the program
End
End Sub

Private Sub cmdkillpercentage_Click()
'This subroutine calculates the kill percentage of every player on the
'team, and then shows all of the kill percentages in the picture box.
picresults.Cls
Dim pos As Single
Dim killpercentage(1 To 30) As Single
Do Until pos >= ctr
    pos = pos + 1
    killpercentage(pos) = (kills(pos) - errors(pos)) / attempts(pos)
    picresults.Print names(pos); " : "; Tab(10); FormatNumber(killpercentage(pos), 3)
Loop
End Sub



Private Sub cmdService_Click()
'This subroutine finds and displays all the statistics involving
'serving for every player on the team
picresults.Cls
Dim pos As Single
    Do Until pos >= ctr
        pos = pos + 1
        picresults.Print names(pos); " : "; Tab(10); aces(pos); " Aces and "; se(pos); " errors."
    Loop
End Sub

Private Sub cmdSwitch_Click()
'goes to the find player screen
frmStats.Hide
frmfind.Show
End Sub

