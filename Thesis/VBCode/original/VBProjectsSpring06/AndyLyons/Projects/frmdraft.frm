VERSION 5.00
Begin VB.Form frmdraft 
   Caption         =   "Draft 1st Round Players"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdclear2 
      Caption         =   "Clear"
      Height          =   735
      Left            =   7320
      TabIndex        =   7
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   855
      Left            =   3480
      TabIndex        =   6
      Top             =   7200
      Width           =   1815
   End
   Begin VB.PictureBox picDisplay2 
      Height          =   6375
      Left            =   6360
      ScaleHeight     =   6315
      ScaleWidth      =   4035
      TabIndex        =   5
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdsimulate 
      Caption         =   "Make Your Choice!"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdorder 
      Caption         =   "Load Draft Order"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtchoice 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.PictureBox picDisplay 
      Height          =   6375
      Left            =   2760
      ScaleHeight     =   6315
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblteam 
      Caption         =   "Enter team draft number"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmdraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pos As Integer
Dim D, X As Single
Dim frmWelcome As Integer
Dim frmFirst As Single
Dim TeamNumber As Integer
Dim PlayerName As String
'2006 NFL Draft Simulator (Draft.vbp)
'frmdraft(frmdraft.frm)
'Andy Lyons
'March 24, 2006
'This form brings the user to the NFL Draft menu. Here the user is able to view the future NFL Draft order. They are then allowed to choose a team and a player they would like to draft.

'clears contents in picturebox
Private Sub cmdclear_Click()
    picDisplay.Cls
End Sub
'clears contents in picturebox
Private Sub cmdclear2_Click()
    picDisplay2.Cls
End Sub

Private Sub cmdorder_Click()
'opens file with 2006 draft order as of 3/22/06
    Open App.Path & "\draftorder.txt" For Input As #2
    Pos = 0
    Do While Not EOF(2)
        Pos = Pos + 1
        Input #2, Rank(Pos), Team(Pos)
    Loop
    TeamSize = Pos
    Close #2
    For Pos = 1 To TeamSize
        picDisplay.Print Rank(Pos), Team(Pos)
    Next Pos
    Open App.Path & "\2006draft.txt" For Input As #3
    Pos = 0
    Do Until EOF(3)
        Pos = Pos + 1
        Input #3, Number(Pos), Player(Pos)
    Loop
    PlayerSize = Pos
    Close #3
End Sub
'returns user to main menu
Private Sub cmdreturn_Click()
    frmNFLDraft.Show
    frmdraft.Hide
End Sub

'clicking this button allows the user to choose a team and make their draft choice. Once choice is made the computer fills in the rest of the choices.
Private Sub cmdsimulate_Click()
    Dim Match, Pos2, Pos As Integer
    TeamNumber = txtchoice.Text
    For Pos = 1 To TeamSize
        If TeamNumber = Rank(Pos) Then
            Match = Pos
        End If
    Next Pos
    
    
    PlayerName = InputBox("Enter a player name for " & Team(Match), "INPUT")
        
Pos2 = 0
    For Pos = 1 To PlayerSize
        Pos2 = Pos2 + 1
        If Pos <> Match Then
            If InStr(Player(Pos), PlayerName) = 0 Then
                picDisplay2.Print Team(Pos); Tab(30); Player(Pos2)
            Else
                Pos2 = Pos2 + 1
                picDisplay2.Print Team(Pos); Tab(30); Player(Pos2)
            End If
        Else
            picDisplay2.Print Team(Pos); Tab(30); PlayerName
        End If
     Next Pos

End Sub



