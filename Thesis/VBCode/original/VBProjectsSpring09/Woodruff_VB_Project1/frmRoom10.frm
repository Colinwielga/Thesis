VERSION 5.00
Begin VB.Form frmRoom10 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSpeak 
      BackColor       =   &H80000015&
      Caption         =   "Speak to glowing thing."
      Height          =   800
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5520
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdCut 
      BackColor       =   &H80000015&
      Caption         =   "Cut Rope?"
      Height          =   800
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdLight 
      BackColor       =   &H80000015&
      Caption         =   "Light Torch!"
      Height          =   800
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdPanic 
      BackColor       =   &H80000015&
      Caption         =   "Panic"
      Height          =   800
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   2500
   End
   Begin VB.CommandButton cmdSecret 
      BackColor       =   &H80000015&
      Caption         =   "Run to Glowing Thing"
      Height          =   800
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   2500
   End
   Begin VB.CommandButton cmdRope 
      BackColor       =   &H80000015&
      Caption         =   "Run to Rope"
      Height          =   800
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Width           =   2500
   End
   Begin VB.CommandButton cmdTorch 
      BackColor       =   &H80000015&
      Caption         =   "Run to Torch"
      Height          =   800
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   2500
   End
   Begin VB.PictureBox pic3D 
      Height          =   2200
      Left            =   8760
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   11
      Top             =   7080
      Width           =   2200
   End
   Begin VB.PictureBox pic2D 
      Height          =   2200
      Left            =   6360
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   10
      Top             =   7080
      Width           =   2200
   End
   Begin VB.PictureBox pic1D 
      Height          =   2200
      Left            =   3960
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   9
      Top             =   7080
      Width           =   2200
   End
   Begin VB.PictureBox pic3C 
      Height          =   2200
      Left            =   8760
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   8
      Top             =   4800
      Width           =   2200
   End
   Begin VB.PictureBox pic2C 
      Height          =   2200
      Left            =   6360
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   7
      Top             =   4800
      Width           =   2200
   End
   Begin VB.PictureBox pic1C 
      Height          =   2200
      Left            =   3960
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   6
      Top             =   4800
      Width           =   2200
   End
   Begin VB.PictureBox pic3B 
      Height          =   2200
      Left            =   8760
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   5
      Top             =   2520
      Width           =   2200
   End
   Begin VB.PictureBox pic2B 
      Height          =   2200
      Left            =   6360
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   4
      Top             =   2520
      Width           =   2200
   End
   Begin VB.PictureBox pic1B 
      Height          =   2200
      Left            =   3960
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   3
      Top             =   2520
      Width           =   2200
   End
   Begin VB.PictureBox pic3A 
      Height          =   2200
      Left            =   8760
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   2
      Top             =   240
      Width           =   2200
   End
   Begin VB.PictureBox pic2A 
      Height          =   2200
      Left            =   6360
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   1
      Top             =   240
      Width           =   2200
   End
   Begin VB.PictureBox pic1A 
      Height          =   2200
      Left            =   3960
      ScaleHeight     =   2145
      ScaleWidth      =   2145
      TabIndex        =   0
      Top             =   240
      Width           =   2200
   End
   Begin VB.Label lblAction 
      BackColor       =   &H80000012&
      Caption         =   "Movement:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000017&
      Caption         =   "Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   12120
      TabIndex        =   19
      Top             =   2520
      Width           =   2535
   End
End
Attribute VB_Name = "frmRoom10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom10
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is the final room in the game.  It is where the user fights a dragon.

Option Explicit
Dim Torch As Boolean
Dim Talk As String
Dim Speak As String
Dim SwordGlow As Boolean
Dim A As Boolean, B As Boolean, C As Boolean
Dim Torchlit As Boolean


Private Sub cmdWin_Click()

    frmRoom10.Visible = False
    frmWin.Visible = True
    
End Sub

Private Sub cmdCut_Click()

    'If SwordGlow = True, user can cut the rope and win
    If SwordGlow = True Then
        MsgBox "You cut the rope and the chandelier falls and misses the dragon completely.  However, he realizes the error of his ways and gives you 100 coins.  Congratulations!", , ""
        frmRoom10.Visible = False
        frmWin.Visible = True
    Else
        MsgBox "Your sword can't cut the rope.  The dragon swipes at you and you lose 2 life.", , ""
        Life = Life - 2
    End If
    
    'Dragon can kill user
    If Life = 0 Then
        frmRoom10.Visible = False
        frmGameOver.Visible = True
    End If
    
End Sub

Private Sub cmdLight_Click()

    'Loads pictures
    pic1A.Picture = LoadPicture(App.Path & "\BossRope.bmp")
    pic2A.Picture = LoadPicture(App.Path & "\BossBoss.bmp")
        If Torchlit = True Then
            pic3A.Picture = LoadPicture(App.Path & "\BossTorchlitYou.bmp")
        Else
            pic3A.Picture = LoadPicture(App.Path & "\BossYouTorch.bmp")
        End If
    pic1B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1C.Picture = LoadPicture(App.Path & "\BossSecret.bmp")
    pic2C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    
    'User given hint
    MsgBox "The room lights up.  There is a large chandelier overhead held up by that rope.  It looks too thick to cut.", , ""
    
    Torchlit = True
    
    pic3A.Picture = LoadPicture(App.Path & "\BossTorchlitYou.bmp")
    
    'Dragon location set
    A = True
    B = False
    C = False
    
    'If dragon kills user
    If Life = 0 Then
        frmRoom10.Visible = False
        frmGameOver.Visible = True
    End If
    
End Sub



Private Sub cmdPanic_Click()

    'The user shouldn't panic, so they are reminded not to
    MsgBox "Don't panic.", , ""
    
End Sub

Private Sub cmdRope_Click()
    
    'Set dragon location
    If A = True Then
        MsgBox "You run by the dragon.  He get's a swipe at you and you lose 1 life.", , ""
        Life = Life - 1
    End If
    
    If C = True Then
        MsgBox "Well, that was stupid.  The dragon was standing right there.  You lose 1 life.  The dragon falls over laughing at you.", , ""
        Life = Life - 1
    End If
    
    'Loads pics and moves user
    pic1A.Picture = LoadPicture(App.Path & "\BossYouRope.bmp")
    pic2A.Picture = LoadPicture(App.Path & "\BossBoss.bmp")
        If Torchlit = False Then
            pic3A.Picture = LoadPicture(App.Path & "\BossTorch.bmp")
        Else
            pic3A.Picture = LoadPicture(App.Path & "\BossTorchlit.bmp")
        End If
    pic1B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1C.Picture = LoadPicture(App.Path & "\BossSecret.bmp")
    pic2C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    
    'Allows user to select 'cut'
    cmdCut.Visible = True
    
    'User can only cut if SwordGlow = true
    If SwordGlow = True Then
        MsgBox "Cut the Rope!", , ""
        
    Else
        MsgBox "The rope is too thick.  Your sword won't cut it.", , ""
    End If
    
    'Shows visible options
    cmdTorch.Visible = True
    cmdRope.Visible = False
    cmdSecret.Visible = True
        
    cmdLight.Visible = False
    cmdSpeak.Visible = False

    'Sets dragon location
    A = True
    B = False
    C = False
    
    'Dragon can kill user
    If Life = 0 Then
        frmRoom10.Visible = False
        frmGameOver.Visible = True
    End If
    
End Sub

Private Sub cmdSecret_Click()

    'If user passes dragon, they lose life
    If B = True Then
        MsgBox "You run by the dragon.  He get's a swipe at you and you lose 1 life.", , ""
        Life = Life - 1
    End If
    
    'Load pictures
    pic1A.Picture = LoadPicture(App.Path & "\BossRope.bmp")
    pic2A.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
        If Torchlit = False Then
            pic3A.Picture = LoadPicture(App.Path & "\BossTorch.bmp")
        Else
            pic3A.Picture = LoadPicture(App.Path & "\BossTorchlit.bmp")
        End If
    pic1B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2B.Picture = LoadPicture(App.Path & "\BossBoss.bmp")
    pic3B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1C.Picture = LoadPicture(App.Path & "\BossYouSecret.bmp")
    pic2C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    
    'User is given the option to say the magic words
    Speak = InputBox("It looks like you should say something to the glowing thing.", "")
    
    'If correct, user can cut rope
    If Speak = SecretCode Then
        MsgBox "The witch's magic words work!  Your sword is glowing.", , ""
        SwordGlow = True
    Else
        MsgBox "Those aren't the words the witch told you.  The dragon swipes at you.  You lose 1 life.", , ""
        Life = Life - 1
        cmdSpeak.Visible = True
    End If
    
    'Shows visible options
    cmdTorch.Visible = True
    cmdRope.Visible = True
    cmdSecret.Visible = False
    
    cmdLight.Visible = False
    cmdSpeak.Visible = True
    cmdCut.Visible = False
    
    'Sets dragon location
    A = False
    B = True
    C = False
    
    'If dragon kills user
    If Life = 0 Then
        frmRoom10.Visible = False
        frmGameOver.Visible = True
    End If
    
End Sub

Private Sub cmdSpeak_Click()

    'User has second option to speak, if he or she failed previously
    Speak = InputBox("It looks like you should say something to the glowing thing.", "")
    
    'Decides if user uses correct words
    If Speak = SecretCode Then
        MsgBox "The witch's magic words work!  Your sword is glowing.", , ""
        SwordGlow = True
    Else
        MsgBox "Those aren't the words the witch told you.  The dragon swipes at you.  You lose 1 life.", , ""
        Life = Life - 1
    End If
    
    'Dragon can kill user here
    If Life = 0 Then
        frmRoom10.Visible = False
        frmGameOver.Visible = True
    End If
    
End Sub

Private Sub cmdTorch_Click()
    
    'Moves user to torch (A3)
    'If user crosses dragon when it's at point 'A' user loses life
    If A = True Then
        MsgBox "You run by the dragon.  He get's a swipe at you and you lose 1 life.", , ""
        Life = Life - 1
    End If
    
    'If user crosses dragon at 'B' user loses life
    If B = True Then
        MsgBox "You run by the dragon.  He get's a swipe at you and you lose 1 life.", , ""
        Life = Life - 1
    End If
    
    'Moves user's picture and dragon's picture
    pic1A.Picture = LoadPicture(App.Path & "\BossRope.bmp")
    pic2A.Picture = LoadPicture(App.Path & "\BossBoss.bmp")
        If Torchlit = False Then
            pic3A.Picture = LoadPicture(App.Path & "\BossYouTorch.bmp")
        Else
            pic3A.Picture = LoadPicture(App.Path & "\BossTorchlitYou.bmp")
        End If
    pic1B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1C.Picture = LoadPicture(App.Path & "\BossSecret.bmp")
    pic2C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    
    'Warns user
    MsgBox "The dragon is moving toward you, run!", , ""
    
    'Hides unusable buttons
    cmdTorch.Visible = False
    cmdRope.Visible = True
    cmdSecret.Visible = True
    
    cmdLight.Visible = True
    cmdSpeak.Visible = False
    cmdCut.Visible = False
    
    'Sets where dragon is so user loses life if he/she crosses dragon
    A = True
    B = False
    C = False
    
    'If dragon 'kills' user this ends the game
    If Life = 0 Then
        frmRoom10.Visible = False
        frmGameOver.Visible = True
    End If
    
End Sub

Private Sub Form_Load()

    
    'Loads pictures
    pic1A.Picture = LoadPicture(App.Path & "\BossbossRope.bmp")
    pic2A.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
        If Torchlit = False Then
            pic3A.Picture = LoadPicture(App.Path & "\BossTorch.bmp")
        Else
            pic3A.Picture = LoadPicture(App.Path & "\BossTorchlit.bmp")
        End If
    pic1B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3B.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1C.Picture = LoadPicture(App.Path & "\BossSecret.bmp")
    pic2C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3C.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic1D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic2D.Picture = LoadPicture(App.Path & "\BossBlank.bmp")
    pic3D.Picture = LoadPicture(App.Path & "\BossYou.bmp")

    'Sets original dragon location
    A = False
    B = False
    C = True
    
End Sub
