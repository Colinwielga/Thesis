VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H80000007&
   Caption         =   "Awakenings"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   2205
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleMode       =   0  'User
   ScaleWidth      =   11070
   Begin VB.PictureBox picTure 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   1200
      Picture         =   "frmIntro.frx":0000
      ScaleHeight     =   5895
      ScaleWidth      =   8535
      TabIndex        =   7
      Top             =   240
      Width           =   8535
   End
   Begin VB.CommandButton cmdEnterDoor 
      Caption         =   "Enter the Door"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Pause and Listen"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdDoor 
      Caption         =   "The Door"
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdLook 
      Caption         =   "Look"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdYell 
      Caption         =   "Yell"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdTextNext 
      Caption         =   "Next"
      Height          =   615
      Left            =   9000
      TabIndex        =   1
      Top             =   6240
      Width           =   855
   End
   Begin VB.PictureBox picIntroTxt 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      ScaleHeight     =   915
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   6960
      Width           =   9015
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim doorcheck As Boolean
Dim text As Integer



Private Sub cmdDoor_Click()
picIntroTxt.Cls
Dim msganswer As Integer

msganswer = MsgBox("Will your approach the door?", vbYesNo) 'create yes/no message box

If msganswer = vbYes Then
    picIntroTxt.Print "You approach the door. You hold your ear to it and listen closely. Only silence."
    picIntroTxt.Print "'Now or never,' you say and grab hold of the handle. You turn it and push."
    picIntroTxt.Print "The door creaks open."
    
    cmdDoor.Visible = False
    cmdEnterDoor.Visible = True
    
    Else
    picIntroTxt.Print "You sit back down. Your body gets colder, your strength diminishes, your hope fades."
    picIntroTxt.Print " "
    picIntroTxt.Print "The darkness closes in on you."
End If

End Sub

Private Sub cmdEnterDoor_Click()
    frmHub.Show               'Changes Form, enter the hub room.
    frmIntro.Hide
End Sub

Private Sub cmdListen_Click()
picIntroTxt.Cls
picIntroTxt.Print "You pause for a moment and listen closely. The faint sound of dripping water can be heard."
picIntroTxt.Print "A small creature scurries through the dark behind you. Are you really alone down here?"

cmdListen.Visible = False 'Make button disappear

If cmdListen.Visible = False And cmdLook.Visible = False And cmdYell.Visible = False And doorcheck = True Then 'After all actions are done, show next action
    cmdDoor.Visible = True
End If

End Sub

Private Sub cmdLook_Click()
picIntroTxt.Cls
picIntroTxt.Print "Your surroundings are covered in a veil of darkness."
picIntroTxt.Print "Even though your eyes have adjusted to the lighting, you still can not see anything except"
picIntroTxt.Print "a single source of light in the distance coming from behind a doorway."

cmdLook.Visible = False     'Make button disappear
doorcheck = True

If cmdListen.Visible = False And cmdLook.Visible = False And cmdYell.Visible = False And doorcheck = True Then 'After all actions are done, show next action
    cmdDoor.Visible = True
End If

End Sub

Private Sub cmdTextNext_Click()
picIntroTxt.Cls

Select Case text            'Cases allow text to change with each button press.
    Case 1
    picIntroTxt.Print "It is a dank, dark, forboding place."
    picIntroTxt.Print "You sit in a circle of overwhelming light casted down from a hole in the ceiling and are"
    picIntroTxt.Print "surrounded on all sides by a never ending darkness."
    text = 2
    
    Case 2
    picIntroTxt.Print "'Hello?' you say to pitch black and it is returned moments later. Your eyes light up but are"
    picIntroTxt.Print "quickly extinguished as you realize it is only your echo."
    text = 3
    
    Case 3
    picIntroTxt.Print "You slowly pick yourself up, your legs trembling with weakness. "
    picIntroTxt.Print "'How long have I been here,' you question yourself. You try to think of how you got here."
    picIntroTxt.Print "Had you fallen down through a hole in the ceiling?"
    text = 4
    
    Case 4
    
    picIntroTxt.Print "But you draw a blank. You can't remember where you were or what you were doing before."
    picIntroTxt.Print "You can't even remember your name. It is all lost in a mental haze. You frantically "
    picIntroTxt.Print "search yourself for something, anything to remind you of who you are."
    text = 5
    
    Case 5
    picIntroTxt.Print "You check your pant's pocket and find a wallet. You open it and see your driver's license."
    text = 6
    
    Case 6
    picIntroTxt.Cls
    
    FirstName = InputBox("Please enter your first name")
        Do Until FirstName <> ""
             FirstName = InputBox("You must enter your first name")
        Loop
    LastName = InputBox("Please enter your last name")
        Do Until LastName <> ""
             LastName = InputBox("You must enter your last name")
        Loop

    picIntroTxt.Cls
    picIntroTxt.Print "Your license says: " & FirstName & " " & LastName & ". Reading it brings back a surge of memories of your identity but it also"
    picIntroTxt.Print "brings back your immobilizing headache. You quickly stop trying to remember which brings. you relief."
    picIntroTxt.Print "Your past is still lost in the thick haze of your mind."
    
    text = 7
    
    Case 7
    picIntroTxt.Print "You place your wallet back into your pocket and notice and object in there. You pull it out."
    picIntroTxt.Print "It is a strange key. A skull is carved into the base of it. Did somebody slip it into your"
    picIntroTxt.Print "pocket? Dozens of possibilites jump to mind making you shudder. Is there somebody else here?"
    
    SkullKey = True
    
    text = 8
    Case 8
    cmdTextNext.Visible = False
    cmdListen.Visible = True
    cmdYell.Visible = True
    cmdLook.Visible = True
                                                    '
    End Select
    
    
End Sub

Private Sub cmdYell_Click()
    picIntroTxt.Cls
    picIntroTxt.Print "You cry for help with all of your strength. You hear your voice reverberate again and again off "
    picIntroTxt.Print "distant walls. 'How big is this place?' You think to yourself. A twinge of fear runs "
    picIntroTxt.Print "through your body. Can you get out?"
    cmdYell.Visible = False 'Make button disappear

    If cmdListen.Visible = False And cmdLook.Visible = False And cmdYell.Visible = False And doorcheck = True Then 'After all actions are done, show next action
        cmdDoor.Visible = True
    End If
    
    
End Sub

Private Sub Form_activate()
    '(Re)Setting all Variables (Global/Public ones too) needed for program to operate (Especially during Form Transition)
    text = 1
    cmdListen.Visible = False
    cmdYell.Visible = False
    cmdLook.Visible = False
    cmdDoor.Visible = False
    doorcheck = False
    cmdEnterDoor.Visible = False
    HubFirst = True
    SkullKey = False
    LeftDoor = False
    Gun = False
    LeftRoomCheck = False
    LockBroken = False
    labcheck = False
    

    picIntroTxt.Print "Slowly, you come to your senses."
    picIntroTxt.Print "You grab your forehead in agony as a bolt of pain strikes your temples."
    picIntroTxt.Print "After what seeems like an eternity, you sit up and peer around."
    End Sub

