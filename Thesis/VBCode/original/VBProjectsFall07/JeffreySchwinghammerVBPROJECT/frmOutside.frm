VERSION 5.00
Begin VB.Form frmOutside 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outside"
   ClientHeight    =   7890
   ClientLeft      =   1365
   ClientTop       =   1755
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   32.875
   ScaleMode       =   0  'User
   ScaleWidth      =   85.125
   Begin VB.CommandButton cmdLiveNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   9000
      TabIndex        =   4
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdKillnext 
      Caption         =   "Next"
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   6720
      Width           =   735
   End
   Begin VB.PictureBox picTure 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   600
      Picture         =   "frmOutside.frx":0000
      ScaleHeight     =   5175
      ScaleMode       =   0  'User
      ScaleWidth      =   409.287
      TabIndex        =   2
      Top             =   240
      Width           =   9255
   End
   Begin VB.CommandButton cmdEndGame 
      Caption         =   "End Game"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin VB.PictureBox picEndtxt 
      Height          =   735
      Left            =   1440
      ScaleHeight     =   675
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   6480
      Width           =   6855
   End
End
Attribute VB_Name = "frmOutside"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txtctr As Integer
Private Sub cmdEndGame_Click()

    'This is Professor Rahal's code.
      ' The file is opened for sequential output, and
      ' data in this case is added to the end of file.
      ' If the file does not exist, it is created.
      Open App.Path & "\Completion.txt" For Append As #1
      Write #1, FirstName, LastName
      Close #1
MsgBox ("Thank you for completing the game, " & FirstName & ". You have been added to the completion list.")

End

End Sub

Private Sub cmdKillnext_Click()
txtctr = txtctr + 1
picEndtxt.Cls
Select Case txtctr
        Case Is = 1
        picEndtxt.Print "You take a look from the cliff edge and see a city off far in the distance."
        picEndtxt.Print "Going there is your best chance of survival. But the distance is far and your"
        picEndtxt.Print "body is weak and battered from escaping the bunker."
        Case Is = 2
        picEndtxt.Print "You still can't remember too much of your identity and it scares you to "
        picEndtxt.Print "think that you may not be able to remember who you are. Maybe you will "
        picEndtxt.Print "find something or meet somebody that trigger the return of your memories."
        Case Is = 3
        picEndtxt.Print "Until then, you are just going to have to get use to living with a stranger"
        picEndtxt.Print "living in haze of past experiences and broken memories. Some have returned"
        picEndtxt.Print "so maybe the rest will come."
        Case Is = 4
        picEndtxt.Print "You notice that you are still caring the pistol. Only one bullet left."
        picEndtxt.Print "You consider tossing it aside when you hear something loud in the distance"
        Case Is = 5
        picEndtxt.Print "You look over the edge of the cliff and see.... your blood runs cold..."
        picEndtxt.Print ""
        picEndtxt.Print "A herd of monsters like the one in the bunker traveling across the sand."
        Case Is = 6
        picEndtxt.Print "'What devil has created these monsters?' you say with terror."
        picEndtxt.Print "Maybe the city isn't the safest place after all."
        Case Is = 7
        picEndtxt.Print "The worse has yet to come..."
        Case Is = 8
        cmdKillnext.Visible = False
        cmdEndGame.Visible = True
        
End Select
End Sub

Private Sub cmdLiveNext_Click()
txtctr = txtctr + 1
picEndtxt.Cls
Select Case txtctr
    Case Is = 1
        picEndtxt.Print "You take a look from the cliff edge and see a city off far in the distance."
        picEndtxt.Print "Going there is your best chance of survival. But the distance is far and your"
        picEndtxt.Print "body is weak and battered from escaping the bunker."
    Case Is = 2
        picEndtxt.Print "The scenery triggers more memories to come back to you. You remember the"
        picEndtxt.Print "world and its dreadful wars. The genetic viral warfare that has been"
        picEndtxt.Print "initiated by... you..."
    Case Is = 3
        picEndtxt.Print "Yes... you. You remember how your efforts to save your wife was overrided"
        picEndtxt.Print "by the people funding it. By the Military. You can't imagine the havoc"
        picEndtxt.Print "the virus has caused."
    Case Is = 4
        picEndtxt.Print "You reach into your pocket and pull out the world's last hope. The vials you have"
        picEndtxt.Print "retreived from the bunker. This is the reverse of the virus, this will undo the"
        picEndtxt.Print "mutations. But it must have another side effect, you realize."
    Case Is = 5
        picEndtxt.Print "It causes temporary amnesia. You must have injected yourself during the"
        picEndtxt.Print "military strike on the bunker and then lost your memory putting your mind"
        picEndtxt.Print "into a haze. The bunker's defense system of theme-designed rooms and puzzles must"
    Case Is = 6
        picEndtxt.Print "have reset while you were unconscious."
        picEndtxt.Print "Your eyes jump to towards the edge of the cliff, you hear something."
        picEndtxt.Print " Something large. You step closer to the edge."
    Case Is = 7
        picEndtxt.Print "Your blood runs cold at the sight."
        picEndtxt.Print "A herd of monsters, just like the one your spouse had mutated into,"
        picEndtxt.Print "are traveling together across the sand in the distance."
    Case Is = 8
        picEndtxt.Print "Your spouse... You will not let her die in vain, you think."
        picEndtxt.Print "You look down at the vial in your hand again. You have a long, difficult road "
        picEndtxt.Print "ahead of you but you will stop the virus mutants and restore them back into people."
    Case Is = 9
        picEndtxt.Print "The worst of your challenges has yet to come..."
        cmdLiveNext.Visible = False
        cmdEndGame.Visible = True
        
    
End Select
End Sub

Private Sub Form_activate()
cmdKillnext.Visible = False
cmdLiveNext.Visible = False
cmdEndGame.Visible = False
txtctr = 0
picEndtxt.Cls
picEndtxt.Print "You step out into burning sun. You are surrounded on all sides by the desert."
picEndtxt.Print "The door you exited through sealed you off from reentering the bunker."
picEndtxt.Print "It is probably for the best. The virus is sealed inside."

If MonsterKill = True Then
    cmdKillnext.Visible = True

Else
    cmdLiveNext.Visible = True
End If

End Sub

