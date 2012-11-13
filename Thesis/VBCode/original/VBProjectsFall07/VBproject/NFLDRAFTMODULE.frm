VERSION 5.00
Begin VB.Form Draft 
   BackColor       =   &H00008000&
   Caption         =   "Draft"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form2"
   ScaleHeight     =   7965
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRoger 
      Height          =   5295
      Left            =   9960
      ScaleHeight     =   5235
      ScaleWidth      =   5955
      TabIndex        =   6
      Top             =   240
      Width           =   6015
   End
   Begin VB.PictureBox picDraftResult 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   4320
      ScaleHeight     =   7995
      ScaleWidth      =   10155
      TabIndex        =   5
      Top             =   6000
      Width           =   10215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   5160
      Width           =   2175
   End
   Begin VB.PictureBox picResult 
      Height          =   5295
      Left            =   2760
      ScaleHeight     =   5235
      ScaleWidth      =   6675
      TabIndex        =   3
      Top             =   240
      Width           =   6735
   End
   Begin VB.CommandButton cmdPostDraft 
      Caption         =   "Push To Move On To Post Draft"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlayerSelect 
      Caption         =   "Please Push To Select Your Desired Player"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdTeam 
      Caption         =   "Please Enter The Team You Are The GM For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Draft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NFL Draft by Justin Buysse and Pete Larson. (NFLPOSTDRAFT.vbp)
'November 6th, 2007
'Form Objective: This form allows the teams to select their desired player
    'depending on if that player has been selected or not.  The user can
    'follow our predetermined draft order or go with their own but each team
    'may only select once.  This form also has added commentary from NFL
    'Commissioner Roger Goodell.
Option Explicit
Dim SPlayer(1 To 32) As String
Dim CTR As Integer
Dim SPlayers As String
Dim Pos As Integer
Dim Found As Boolean
Dim Found2 As Boolean
Dim Pos1 As Integer
Dim CTR1 As Integer
Dim Teams As String
Dim DraftCtr As Integer

Private Sub cmdPlayerSelect_Click()
'This button allows the user to draft the player.  The program will be able to tell if the
'player is still available in the draft.  If the player is available; the team name, player's name
'and draft pick number will be displayed.  If the player is unavailable, the player's name and the
'team that has already selected him will be displayed.
'If the player is not in the draft, an error message will be displayed
    cmdPlayerSelect.Enabled = True
    cmdTeam.Enabled = True
    cmdPostDraft.Enabled = True
    Dim Athlete As String
    Athlete = InputBox("Please select your player", "Selecting A Player", , 1000, 1000)
    Pos = 0
    Found = False
    Do While (Found = False And Pos < CTR_Players)
        Pos = Pos + 1
        If LCase(Player(Pos)) = LCase(Athlete) Then
            If Selected(Pos) = "NONE" Then
                Found = True
            Else
                Found = True
                MsgBox Athlete & " has already been selected by " & Selected(Pos), , "Selecting A Player"
            End If
        End If
    Loop
    If Found = True And Selected(Pos) = "NONE" Then
        Selected(Pos) = Teams
        picRoger.Picture = LoadPicture(App.Path & "\RogerGoodell.jpg")
        MsgBox ("Roger Goodell: The " & Teams & " have selected " & Athlete & " with the number " & DraftCtr & " overall pick."), , "Selecting A Player"
        picDraftResult.Print "The " & Teams & " selected " & Athlete & " with the overall number " & DraftCtr & " overall pick."
    ElseIf Found = False Then
        MsgBox (Athlete & " is not in the draft"), , "Selecting A Player"
    End If
End Sub

Private Sub cmdPostDraft_Click()
'This button will hide the PreDraft and Draft forms and allow the user to move on to the Post-Draft form
'The Post-Draft form is the last form of the project.
    PreDraft.Hide
    Draft.Hide
    PostDraft.Show
End Sub

Private Sub cmdQuit_Click()
'This button allows the user to quit the program at any time.
    End
End Sub

Private Sub cmdTeam_Click()
'This button reads the teams from the array.  It will ask the user what team they represent and once this question is answered,
'their team logo will appear in the picture box.
'If the team that is inputted is not a valid team, an error message will appear stating as such.
'The teams are all case sensitive so as long as the spelling is correct, the correct team logo will appear.
    cmdTeam.Enabled = True
    cmdPlayerSelect.Enabled = True
    cmdPostDraft.Enabled = False
    Dim Team(1 To 32) As String
    CTR_Teams = 0
    Open App.Path & "\TeamsRanks.txt" For Input As #1
    Do Until EOF(1)
        CTR_Teams = CTR_Teams + 1
        Input #1, Team(CTR_Teams), Rank(CTR_Teams)
    Loop
    Close #1
    Teams = InputBox("Please enter the team you are the GM for", "Input Team", , 1000, 1000)
    If LCase(Teams) = LCase(Team(1)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Buffalo.gif")
    ElseIf LCase(Teams) = LCase(Team(2)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Miami.gif")
    ElseIf LCase(Teams) = LCase(Team(3)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\New%20England.gif")
    ElseIf LCase(Teams) = LCase(Team(4)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\New%20York%20Jets.gif")
    ElseIf LCase(Teams) = LCase(Team(5)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Baltimore.gif")
    ElseIf LCase(Teams) = LCase(Team(6)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Cincinnati.gif")
    ElseIf LCase(Teams) = LCase(Team(7)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Cleveland.gif")
    ElseIf LCase(Teams) = LCase(Team(8)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Pittsburgh.gif")
    ElseIf LCase(Teams) = LCase(Team(9)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Houston.gif")
    ElseIf LCase(Teams) = LCase(Team(10)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Indianapolis.gif")
    ElseIf LCase(Teams) = LCase(Team(11)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Jacksonville.gif")
    ElseIf LCase(Teams) = LCase(Team(12)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Tennessee.gif")
    ElseIf LCase(Teams) = LCase(Team(13)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Denver.gif")
    ElseIf LCase(Teams) = LCase(Team(14)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Kansas%20City.gif")
    ElseIf LCase(Teams) = LCase(Team(15)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Oakland.gif")
    ElseIf LCase(Teams) = LCase(Team(16)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\San%20Diego.gif")
    ElseIf LCase(Teams) = LCase(Team(17)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Dallas.gif")
    ElseIf LCase(Teams) = LCase(Team(18)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\New%20York%20Giants.gif")
    ElseIf LCase(Teams) = LCase(Team(19)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Philadelphia.gif")
    ElseIf LCase(Teams) = LCase(Team(20)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Washington.gif")
    ElseIf LCase(Teams) = LCase(Team(21)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Chicago.gif")
    ElseIf LCase(Teams) = LCase(Team(22)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Detroit.gif")
    ElseIf LCase(Teams) = LCase(Team(23)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Green%20Bay.gif")
    ElseIf LCase(Teams) = LCase(Team(24)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Minnesota.gif")
    ElseIf LCase(Teams) = LCase(Team(25)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Atlanta.gif")
    ElseIf LCase(Teams) = LCase(Team(26)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Carolina.gif")
    ElseIf LCase(Teams) = LCase(Team(27)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\New%20Orleans.gif")
    ElseIf LCase(Teams) = LCase(Team(28)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Tampa%20Bay.gif")
    ElseIf LCase(Teams) = LCase(Team(29)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Arizona.gif")
    ElseIf LCase(Teams) = LCase(Team(30)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\St_%20Louis.gif")
    ElseIf LCase(Teams) = LCase(Team(31)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\San%20Francisco.gif")
    ElseIf LCase(Teams) = LCase(Team(32)) Then
        DraftCtr = DraftCtr + 1
        picResult.Picture = LoadPicture(App.Path & "\Seattle.gif")
    Else
        MsgBox ("Error, you have not entered a valid NFL team.  Please choose again."), , Error
    End If
End Sub

