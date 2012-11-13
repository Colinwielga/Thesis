VERSION 5.00
Begin VB.Form frmHub 
   BackColor       =   &H80000007&
   Caption         =   "Hub Room"
   ClientHeight    =   9435
   ClientLeft      =   -105
   ClientTop       =   2055
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   12435
   Begin VB.PictureBox picTure 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   840
      Picture         =   "frmHub.frx":0000
      ScaleHeight     =   6855
      ScaleWidth      =   10695
      TabIndex        =   4
      Top             =   120
      Width           =   10695
   End
   Begin VB.CommandButton cmdBreakLock 
      Caption         =   "Break the Lock"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheckSelect 
      Caption         =   "Click to Observe Selected Object"
      Height          =   615
      Left            =   9960
      TabIndex        =   2
      Top             =   7200
      Width           =   2415
   End
   Begin VB.ListBox lstCheckList 
      Height          =   1425
      ItemData        =   "frmHub.frx":12202
      Left            =   9960
      List            =   "frmHub.frx":12204
      TabIndex        =   1
      Top             =   7920
      Width           =   2415
   End
   Begin VB.PictureBox picHubTxt 
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
      Left            =   600
      ScaleHeight     =   915
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   7920
      Width           =   9015
   End
End
Attribute VB_Name = "frmHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EmblemCheck As Boolean
Dim Skull As Boolean
Dim BoneText As Integer
Dim answer As Integer





Private Sub cmdBreakLock_Click()
    cmdBreakLock.Visible = False
    picHubTxt.Cls
    picHubTxt.Print "You aim your pistol at the lock with your right hand and cover your face"
    picHubTxt.Print "with your left. You pull the trigger and sparks fly everywhere."
    picHubTxt.Print "The broken lock hits the ground with a thud. The door is unlocked."
    LockBroken = True
    
End Sub

Private Sub cmdCheckSelect_Click()
Dim msganswer As Integer ' declaring local variable

    picHubTxt.Cls   'clear the screen for new text
    
    '******************************
    
    If lstchecklist = "Center Door" Then
        If EmblemOne = True And EmblemTwo = True And EmblemThree = True And EmblemCheck = True Then
            answer = MsgBox("You have found THREE pieces of the EMBLEM. Place it into the depression?", vbYesNo, "Emblem")
            If answer = vbYes Then
            MsgBox "The pieces fit perfectly into the depression. The door slides open and you step through."
            frmStairs.Show
            frmHub.Hide
            Else
                picHubTxt.Print "You decided not to put all three pieces into the depression. You are running out of time!"
            'Make Y/N Message Box for the user to use the three emblem pieces.
            
            End If
        ElseIf EmblemOne = True And EmblemTwo = True And EmblemThree = False And EmblemCheck = True Then
                picHubTxt.Print "On the middle door, you notice that there is a round depression on the door."
                picHubTxt.Print "It looks like a Circular Emblem could fit in it. You currently have TWO piece of the emblem."
                picHubTxt.Print "ONE more piece is needed to complete the Emblem. You wonder where it could be..."
        ElseIf EmblemOne = True And EmblemCheck = True Then
                picHubTxt.Print "On the middle door, you notice that there is a round depression on the door."
                picHubTxt.Print "It looks like a circular EMBLEM could fit in it. You currently have ONE piece of the Emblem."
        'This is last because it might throw off the other Ifs
        ElseIf EmblemCheck = True And EmblemOne = False And EmblemTwo = False And EmblemThree = False Then
                picHubTxt.Print "You wander up to the massive door in the center. You notice that there is an imprint"
                picHubTxt.Print "on the door. It looks like a EMBLEM of some sort could fit into it."
                picHubTxt.Print "Maybe you should try finding the Emblem? Is it around here?"
                If BoneText = 2 Then
                    BoneText = 3
                End If
        ElseIf EmblemCheck = False And EmblemOne = False And EmblemTwo = False And EmblemThree = False Then
                picHubTxt.Print "You wander up to the massive door in the center. You notice that there is an imprint"
                picHubTxt.Print "on the door. It looks like a EMBLEM of some sort could fit into it."
                picHubTxt.Print "Maybe you should try finding the Emblem?"
                EmblemCheck = True
                If BoneText = 2 Then
                BoneText = 3 'change BoneCheck for case
                 End If
        End If
        
    End If
    
    '******************************
    If Null Then
        picHubTxt.Print "Please select from the list at left what you want to observe."
    End If
    If lstchecklist = "Pile of Bones" Then
        If EmblemOne = True And Skull = False Then
            picHubTxt.Print "There is a grotesque pile of bones on the floor. It makes you queasy just looking at."
            picHubTxt.Print "You look closer at it again. There is a Skull lying in the pile."
            
            msganswer = MsgBox("Will you take the skull?", vbYesNo) 'create yes/no message box
        
                If msganswer = vbYes Then
                    picHubTxt.Cls
                    picHubTxt.Print "Insects scatter away as you picked up the Skull. You might be able to use it."
                    BoneText = 4
                    
                    Skull = True 'add skull to the item list, this doesn't work
                    'cmbItembox.AddItem "Skull"
                Else
                    picHubTxt.Cls
                    picHubTxt.Print "You decided not to pick up the skull. It would be disgraceful to touch it."
                End If
            
        Else
                Select Case BoneText
                Case Is < 3
                    picHubTxt.Print "There is a grotesque pile of bones on the floor. It makes you queasy just looking at."
                    picHubTxt.Print "Who was this person? Another lost soul?        Is this your fate?"
                    
                    If BoneText = 1 Then
                        BoneText = 2
                    End If
                
                Case 3
                    If LeftDoor = False And EmblemOne = False Then

                        picHubTxt.Print "There is a grotesque pile of bones on the floor. It makes you queasy just looking at."
                        picHubTxt.Print "You look closer in to the pile and see something shiny hiding under bones."
                        
                        msganswer = MsgBox("Will you reach for it the object?", vbYesNo) 'create yes/no message box
        
                        If msganswer = vbYes Then
                            picHubTxt.Cls
                            picHubTxt.Print "You carefully reach for the object, trying not to touch the bones in the process."
                            
                            picHubTxt.Print "You pull out a BROKEN EMBLEM PIECE. This might be useful so you hold onto it."
                        EmblemOne = True
                        'lstItembox.AddItem "Emblem Piece #1"
                        
                        Else
                            picHubTxt.Cls
                            picHubTxt.Print "You decide against it. You are too creeped out to get near the skeleton."
                        End If
                    End If
                Case Is > 3
                        picHubTxt.Cls
                        picHubTxt.Print "There is nothing to earn from searching the bone pile. You decide not to think"
                        picHubTxt.Print "about it anymore."
        
                End Select
            End If
    End If
    
    '******************************
    
    If lstchecklist = "Entrance Door" Then
        picHubTxt.Print "You check the door you came in through. Nope. It is still locked."
    End If
    
    '******************************
    
    If lstchecklist = "Door on the Right" Then
        picHubTxt.Cls
        If EmblemThree = True Then
        picHubTxt.Print "There is no reason to go back to the lab. You need to escape now!"
        Else
            If LockBroken = True Then
                msganswer = MsgBox("Will you go through the door?", vbYesNo)
                
                If msganswer = vbYes Then
                    frmHallway.Show
                    frmHub.Hide
                Else
                picHubTxt.Print "You decided not to go through the door."
                End If
            
            Else
                If Gun = True Then
                    picHubTxt.Print "You go up to the door on the right side. You see a worn out padlock on it."
                    picHubTxt.Print "You think that a gunshot might be able to break to lock."
                    
                    cmdBreakLock.Visible = True
                
                Else
                    picHubTxt.Print "You go up to the door on the right side. You see a worn out padlock on it."
                    picHubTxt.Print "You pound on the door and try to break the lock. It is too strong for you. Maybe"
                    picHubTxt.Print "you can find something to break the padlock."
                End If
            End If
        End If
    End If
    '******************************
    
    If lstchecklist = "Door on the Left" Then
        picHubTxt.Cls
        If LeftRoomCheck = False Then
        
            If LeftDoor = True Then
                msganswer = MsgBox("Will you walk through the door", vbYesNo) 'create yes/no message box
            
                    If msganswer = vbYes Then
                        picHubTxt.Cls
                        picHubTxt.Print "You open the door and go into the next room."
                        frmLibrary.Show               'Changes Form
                        frmHub.Hide
                    Else
                        picHubTxt.Print "You decided not to go through the door."
                    End If
            
            
            Else
                If Skull = False Then
                    picHubTxt.Print "You approach the door on the left. It is locked. Next to the door is a pedastal with a"
                    picHubTxt.Print "a keyhole on it. Over it is a ghastly picture of a SKULL. You try the skull key you found"
                    picHubTxt.Print "in your pocket. It fits but doesn't do anything. Maybe something has to go on the pedastal."
                End If
            
                If Skull = True Then
                    picHubTxt.Print "You approach the door on the left. It is locked. Next to the door is a pedastal with"
                    picHubTxt.Print "a keyhole on it. Over it is a ghastly picture of a SKULL."
                
                    msganswer = MsgBox("Will you place the SKULL on pedastal and use your SKULL KEY", vbYesNo) 'create yes/no message box
            
                    If msganswer = vbYes Then
                        picHubTxt.Cls
                        picHubTxt.Print "You are more than happy to set the SKULL on the pedastal. You don't want to hold it anymore."
                        picHubTxt.Print "You insert the SKULL KEY and turn it. A distinct click sound: the door is unlocked now."
                        LeftDoor = True
                        BoneText = 4
                        'lstItembox.RemoveItem "Skull" this doesn't work
                        
                    Else
                        picHubTxt.Cls
                        picHubTxt.Print "You decided against it."
                    End If
                End If
            End If
        Else
        picHubTxt.Cls
            picHubTxt.Print "You have been locked out of this room again. The Skull key won't work anymore. Which"
            picHubTxt.Print "is okay, you didn't really want to go in there again anyway."
        End If
    End If
    '******************************
    If lstchecklist = "Ceiling" Then
        picHubTxt.Cls
        picHubTxt.Print "You cast your eyes upward at the ceiling. It is dark and covered with spider webs."
        picHubTxt.Print "You see some bats hanging but you don't intend to bother them."
    End If

End Sub



Private Sub Form_activate()

'This one goes first or it messes up the next if then statement
If HubFirst = False Then 'Make sure that when player checks middle door, it will say correct thing. Not sure if necessary at the moment
    EmblemCheck = True
End If

If HubFirst = True Then
    picHubTxt.Cls
    picHubTxt.Print "The door swings shut behind you locking you in this strange room."
    HubFirst = False
    EmblemCheck = False 'flag for Events
    Skull = False   'flag for Events
    
    'Setting to false at initial load for variables to work
    EmblemOne = False
    EmblemTwo = False
    EmblemThree = False
    
End If
'Filling up the list box
 If Gun = False Then
    lstchecklist.AddItem "Entrance Door"
    lstchecklist.AddItem "Pile of Bones"
    lstchecklist.AddItem "Ceiling"
    lstchecklist.AddItem "Door on the Left"
    lstchecklist.AddItem "Door on the Right"
    lstchecklist.AddItem "Center Door"
End If

BoneText = 1 ' Setting Variable/Integer for Skeleton Text
cmdBreakLock.Visible = False

    
End Sub

Private Sub lstItembox_click()
If EmblemOne = True Then
    lstItembox.AddItem "Emblem Piece #1"
    Else
        lstItembox.RemoveItem "Emblem Piece #1"
End If
    
    
If EmblemTwo = True Then
    lstItembox.AddItem "Emblem Piece #2"
Else
    lstItembox.RemoveItem "Emblem Piece #2"
End If

If EmblemThree = True Then
        lstItembox.AddItem "Emblem Piece #3"
    Else
        lstItembox.RemoveItem "Emblem Piece #3"
End If
If SkullKey = True Then
    lstItembox.AddItem "Skull Key"
Else
    lstItembox.RemoveItem "Skull Key"
End If

End Sub

