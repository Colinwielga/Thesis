VERSION 5.00
Begin VB.Form frmRoom3 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000015&
      Caption         =   "Back to entrance"
      Height          =   800
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   2500
   End
   Begin VB.CommandButton cmdAnswer 
      BackColor       =   &H80000015&
      Caption         =   "Answer the troll"
      Height          =   800
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   2500
   End
   Begin VB.PictureBox picRoom3 
      Height          =   6855
      Left            =   2880
      ScaleHeight     =   6795
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   360
      Width           =   8775
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
      TabIndex        =   5
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Movement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblStoryRoom3 
      BackColor       =   &H80000017&
      Caption         =   $"frmRoom3.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1695
      Left            =   2520
      TabIndex        =   1
      Top             =   7440
      Width           =   9615
   End
End
Attribute VB_Name = "frmRoom3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom3
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is a 'room' of the game.  It is where the user talks to a troll.

Option Explicit

Private Sub cmdAnswer_Click()

    'Sets correct answer and gets input box for user to answer
    'If user is wrong he/she loses 1 life
    Dim answer As Single
    
    'If there is an error, the On Error will redirect it to the end
    On Error GoTo InputError
    answer = InputBox("How many?", "")
    
    
        If answer = 6 Then
            frmRoom3.Visible = False
            frmRoom4.Visible = True
                If TrollCoins = False Then
                    MsgBox "Hey!  He is a pretty nice troll.  You got the answer right and 5 coins!  You continue on to the next room.", , ""
                    Coins = Coins + 5
                Else
                    MsgBox "Right.  Moving on.", , ""
                End If
            TrollCoins = True
        Else
            MsgBox "Nope, you're wrong. He beans you over the head with his big troll club.  You lose 1 life and head back to the entrance.", , ""
            Life = Life - 1
            frmRoom3.Visible = False
            frmRoom1.Visible = True
        End If
    
        'If lose of life kills user, game over
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
InputError:

        
End Sub

Private Sub cmdBack_Click()

    'User leaves to room 1
    frmRoom3.Visible = False
    frmRoom1.Visible = True
    
End Sub

Private Sub Form_Load()

    picRoom3.Picture = LoadPicture(App.Path & "\Troll.jpg")
    
End Sub

Private Sub picResults_Click()

End Sub

