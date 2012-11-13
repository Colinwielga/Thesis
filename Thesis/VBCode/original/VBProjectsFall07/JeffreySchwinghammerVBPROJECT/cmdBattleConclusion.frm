VERSION 5.00
Begin VB.Form frmBattleConclusion 
   BackColor       =   &H80000007&
   ClientHeight    =   9135
   ClientLeft      =   540
   ClientTop       =   1890
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   10605
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   120
      Picture         =   "cmdBattleConclusion.frx":0000
      ScaleHeight     =   6495
      ScaleWidth      =   10335
      TabIndex        =   7
      Top             =   240
      Width           =   10335
   End
   Begin VB.CommandButton cmdEscape 
      Caption         =   "Escape"
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdLivetxt 
      Caption         =   "Next"
      Height          =   615
      Left            =   8640
      TabIndex        =   5
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdKillTxt 
      Caption         =   "Next"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   615
      Left            =   9120
      TabIndex        =   3
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton cmdStayBack 
      Caption         =   "Stay Back"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdApproach 
      Caption         =   "Approach It?"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   6960
      Width           =   1575
   End
   Begin VB.PictureBox picconclusiontxt 
      Height          =   975
      Left            =   1680
      ScaleHeight     =   915
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   7800
      Width           =   7215
   End
End
Attribute VB_Name = "frmBattleConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim livectr As Integer
Dim killctr As Integer
Dim ctr As Integer

Dim killanswer As Integer 'do I need this?
Dim answer As Integer
Private Sub cmdApproach_Click()
cmdApproach.Visible = False
cmdStayBack.Visible = False
cmdNext.Visible = True
picconclusiontxt.Cls
picconclusiontxt.Print "You decide that the monster is no longer a threat and you slowly"
picconclusiontxt.Print "walk over to it. The monster gurgles and coughs as its life slips"
picconclusiontxt.Print "away. You can feel an ethereal connection with it... you can sense its pain."
End Sub

Private Sub cmdEscape_Click()
    frmHub.Show
    frmBattleConclusion.Hide
    End Sub

Private Sub cmdKillTxt_Click()
killctr = killctr + 1
Select Case killctr
    Case Is = 1
    picconclusiontxt.Cls
    picconclusiontxt.Print "Over the alarms, you hear a computer voice over the intercom speakers..."
   'picconclusiontxt.Print "'Please Evacuate the Area' it says, 'There is Containment Structural Damage.'"
   ' picconclusiontxt.Print "'Preparing for Lockdown'"
    Case Is = 2
        
        answer = MsgBox("Please Evacuate the Area! There is Containment Structural Damage. Preparing for Lockdown", 48, "InterCom Announcement")
    
    Case Is = 3
    
    picconclusiontxt.Cls
    picconclusiontxt.Print "Your eyes catch the glint of something in the Testing Room from which the monster came"
    picconclusiontxt.Print "from. You slowly peek your head into the room. You feel certain its safe but you approach"
    picconclusiontxt.Print "it carefully. You get close enough to recognize what it is and your eyes light up!"
    
    Case Is = 4
    picconclusiontxt.Cls
    picconclusiontxt.Print "It is the third piece of the emblem! Just what you need to open the last door!"
    picconclusiontxt.Print "As you pocket it, there is another computer announcement"
    EmblemThree = True
    
    Case Is = 5
    answer = MsgBox("Restricted Access Doors have been unlocked to aid in evacuation! Please Evacuate the Building!", 48, "Intercom Announcement")
    
    Case Is = 6
    picconclusiontxt.Print "You hear the click of an unlocking door."
    
    cmdKillTxt.Visible = False
    cmdEscape.Visible = True
    

End Select
End Sub

Private Sub cmdLivetxt_Click()
livectr = livectr + 1
Select Case livectr
Case Is = 1
    answer = MsgBox("Please Evacuate the Area! There is Containment Structural Damage. Preparing for Lockdown", 48, "InterCom Announcement")

Case Is = 2
    picconclusiontxt.Cls
    picconclusiontxt.Print "Despite the Intercom System telling you to get out of this place, you are compelled to get"
    picconclusiontxt.Print "closer to the dying creature. You place you hand on the creature's head. A surge of memories flood"
    picconclusiontxt.Print "your mind. You life, your experiences, your identity is recollected in a moment."
Case Is = 3
    picconclusiontxt.Cls
    picconclusiontxt.Print "You remember the experiments you were running, the biological formulas you were creating,"
    picconclusiontxt.Print "and how it all fell apart. You remember your dying spouse and your attempts to create a cure"
    picconclusiontxt.Print "through a genetic reprogramming virus. You were about to cure your spouse's cancer but..."
Case Is = 4
    picconclusiontxt.Cls
    picconclusiontxt.Print "Something went wrong. Dreadfully wrong."
    picconclusiontxt.Print "Your military contractors wanted to use the virus for warfare purposes and they were impatient."
    picconclusiontxt.Print "They stormed your compound and adjusted the virus and tested it on your spouse."
    
Case Is = 5
    picconclusiontxt.Cls
    picconclusiontxt.Print "An error in the virus programming had resulted in the grotesque transformation of your spouse"
    picconclusiontxt.Print "from human into beast."
    picconclusiontxt.Print "The result lies before you."

Case Is = 6
    picconclusiontxt.Cls
    picconclusiontxt.Print "You come back to reality abruptly, realizing the danger of the situation. The creature moans in "
    picconclusiontxt.Print "pain. Its life is slowly escaping its body. It turns it head weakly in your direction.  You feel an"
    picconclusiontxt.Print "ethereal connection with your spouse. You feel its humanity again."

Case Is = 7
    picconclusiontxt.Cls
    picconclusiontxt.Print "And then that connection fades. It breathes its last."
    picconclusiontxt.Print " "
    picconclusiontxt.Print "It lies lifeless before you. But it managed to communicate something to you..."
Case Is = 8
    picconclusiontxt.Cls
    picconclusiontxt.Print "You understand what you must do. You stand up and move to the Testing Area room where the creature"
    picconclusiontxt.Print "was imprisoned. Inside you find the third piece of the Emblem. The last of three, the key to escaping this place."
    picconclusiontxt.Print "There is also a set of vials. They are labeled: Genetic Recombinator Virus. You hear another announcement:"
    EmblemThree = True
    

Case Is = 9
    answer = MsgBox("Restricted Access Doors have been unlocked to aid in evacuation! Please Evacuate the Building!", 48, "Intercom Announcement")
    
Case Is = 10
    picconclusiontxt.Cls
    picconclusiontxt.Print "You hear the click of an unlocking door. You take the Emblem and the vials into your pocket and leave."
    cmdLivetxt.Visible = False
    cmdEscape.Visible = True
    
End Select
End Sub

Private Sub cmdNext_Click()
ctr = ctr + 1
Select Case ctr
    Case Is = 1
        picconclusiontxt.Cls
        picconclusiontxt.Print "The monster... thing... being could still get up and attack you again you think."
        picconclusiontxt.Print "You feel a twinge of remorse as you imagine the execution act."
    Case Is = 2
        killanswer = MsgBox("Should you kill the monster?", vbYesNo)
        picconclusiontxt.Cls
        If killanswer = vbYes Then
           
            cmdKillTxt.Visible = True
            picconclusiontxt.Print "You check your hand gun's clip: one bullet left. That's all you need, you think."
            picconclusiontxt.Print "You take aim between the monster's temples. The room echoes with the sound of the gunshot and the pulsing alarm."
            picconclusiontxt.Print "You are safe now. The monster has been defeated... "
            MonsterKill = True
        Else
            
            
            cmdLivetxt.Visible = True
            picconclusiontxt.Print "You check your hand gun's clip: one bullet left. This is not neccessary, you think."
            picconclusiontxt.Print "This creature has been through enough. There is no reason to put it through more."
            picconclusiontxt.Print "You look into its eyes and feel as if you can sense its graditude."
            MonsterKill = False
        End If
        cmdNext.Visible = False
End Select
        
End Sub

Private Sub cmdStayBack_Click()
picconclusiontxt.Cls
picconclusiontxt.Print "You decide it is best not to get close to it. You wipe the blood off your face.You look"
picconclusiontxt.Print "around: much of the room has been destroyed in the fight. Broken glass and machinery "
picconclusiontxt.Print "is strewn across the floor. The door is still locked there must be someway to get out."

End Sub

Private Sub Form_activate()
ctr = 0
livectr = 0
killctr = 0

cmdLivetxt.Visible = False
cmdKillTxt.Visible = False
cmdEscape.Visible = False

cmdNext.Visible = False
picconclusiontxt.Cls
picconclusiontxt.Print "You keep your eyes on the fallen creature. Its limbs sprawled out and its chest heaving "
picconclusiontxt.Print "up and down as it chokes for air. It doesn't have much longer to live but"
picconclusiontxt.Print "it could still be dangerous..."
End Sub
