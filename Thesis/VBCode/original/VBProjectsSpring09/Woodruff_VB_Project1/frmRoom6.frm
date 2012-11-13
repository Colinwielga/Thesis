VERSION 5.00
Begin VB.Form frmRoom6 
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
   Begin VB.CommandButton cmdOtherLadder 
      BackColor       =   &H80000015&
      Caption         =   "Try other ladder."
      Height          =   800
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   2500
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000015&
      Caption         =   "Up first ladder."
      Height          =   800
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   2500
   End
   Begin VB.CommandButton cmdWait 
      BackColor       =   &H80000015&
      Caption         =   "Wait"
      Height          =   800
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2500
   End
   Begin VB.CommandButton cmdSnatch 
      BackColor       =   &H80000015&
      Caption         =   "Snatch"
      Height          =   800
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2500
   End
   Begin VB.PictureBox picRoom6 
      Height          =   5295
      Left            =   2880
      ScaleHeight     =   5235
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   840
      Width           =   9015
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
      Left            =   12240
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
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
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label lblStoryRoom6 
      BackColor       =   &H80000017&
      Caption         =   $"frmRoom6.frx":0000
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
      Height          =   2175
      Left            =   2880
      TabIndex        =   1
      Top             =   6480
      Width           =   9015
   End
End
Attribute VB_Name = "frmRoom6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom6
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is a 'room' of the game.  It is where the user tries to steal a key from a snake.

Option Explicit
Dim SnatchCounter As Integer

Private Sub cmdBack_Click()
    
    'Moves user to room 5
    'Resets snatch puzzle
    frmRoom6.Visible = False
    frmRoom5.Visible = True
    SnatchCounter = 0
    
    
End Sub

Private Sub cmdOtherLadder_Click()

    'Moves user to room 2
    'Resets snatch puzzle
    'Makes the use of the Ladder available from room 2
    frmRoom6.Visible = False
    frmRoom2.Visible = True
    Ladder = True
    SnatchCounter = 0
    
    
    
End Sub

Private Sub cmdSnatch_Click()
    
    'In this puzzle, the user has to either snatch or wait depending on
    'if the snake is looking at him/her.
    'If he/she makes a wrong decision, he/she lose 1 life.
    'If he/she is right, then he/she gets a key
    'These are the messages the user gets when he/she attempts to snatch the key
    SnatchCounter = SnatchCounter + 1
    
        If SnatchCounter = 2 Or SnatchCounter = 4 Or SnatchCounter = 7 Then
            MsgBox "You successfully grab the key without the snake noticing.  Plus, got 5 coins that were next to the key!  You better get out of here before the snake sees its key is gone.", , ""
            SwordKey = True
            Coins = Coins + 5
            cmdSnatch.Visible = False
            cmdWait.Visible = False
            
        Else
            Select Case SnatchCounter
                Case 1
                    MsgBox "Crap.  You tried to soon.  The snake bit you and you lost 1 life.", , ""
                    Life = Life - 1
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaAway.bmp")
                Case 3
                    MsgBox "You missed your chance.  You got bit and lose 1 life.  However, it looks like the snake isn't look...", , ""
                    Life = Life - 1
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaAway.bmp")
                Case 5
                    MsgBox "Bad timing.  You get bit.  It looks like you're only going to get a few more chances.", , ""
                    Life = Life - 1
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaYou.bmp")
                Case 6
                    MsgBox "Bad timing.  You get bit.  It looks like you're only going to get a few more chances.", , ""
                    Life = Life - 1
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaAway.bmp ")
                Case 8 To 100
                    MsgBox "You get bit and it looks like the snake's keeping its eyes on you completely.  Better leave and wait for it to calm down.", , ""
                    Life = Life - 1
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaYou.bmp")
            End Select
            
        End If
        
    'If the snake kills the user, it's game over
    If Life = 0 Then
        frmRoom1.Visible = False
        frmRoom2.Visible = False
        frmRoom3.Visible = False
        frmRoom4.Visible = False
        frmRoom5.Visible = False
        frmRoom6.Visible = False
        frmRoom7.Visible = False
        frmRoom8.Visible = False
        frmRoom9.Visible = False
        frmRoom10.Visible = False
        frmGameOver.Visible = True
    End If
    
End Sub

Private Sub cmdWait_Click()

    'These are the what happens and the messages the user gets
    'when he/she waits.
    SnatchCounter = SnatchCounter + 1
    
    
            Select Case SnatchCounter
                Case 1
                    MsgBox "Good, the snake was eyeing you anyway...", , ""
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaAway.bmp")
                Case 2
                    MsgBox "You should have tried for it.  Give it a second and try again.", , ""
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaYou.bmp")
                Case 3
                    MsgBox "The snake is really close to the key.  But you might want to try for it.", , ""
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaAway.bmp")
                Case 4
                    MsgBox "You missed another shot.  Wait two seconds go for it.  You might have only one more chance.", , ""
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaYou.bmp")
                Case 5
                    MsgBox "Wait for it...", , ""
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaYou.bmp")
                Case 6
                    MsgBox "This is it...", , ""
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaAway.bmp")
                Case 7
                    MsgBox "You missed your chance.  The snakes looking right at you.  Leave and try come back later.", , ""
                    picRoom6.Picture = LoadPicture(App.Path & "\kaaYou.bmp")
            End Select
End Sub

Private Sub Form_Load()

    picRoom6.Picture = LoadPicture(App.Path & "\kaaYou.bmp")
    
End Sub

