VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReselect 
      Caption         =   "Reselect Character"
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm Character"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   4800
      Width           =   2175
   End
   Begin VB.PictureBox picCharacter 
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.PictureBox picCharacterStats 
         Height          =   1095
         Left            =   1440
         ScaleHeight     =   1035
         ScaleWidth      =   5355
         TabIndex        =   3
         Top             =   3480
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim C As Single


Private Sub cmdConfirm_Click()
    
    'Confirms the character and starts the game
    frmStart.Visible = False
    frmRoom1.Visible = True
    
End Sub

Private Sub cmdReselect_Click()

    'Allows user to reselect the character
    frmStart.Visible = False
    frmCharacter.Visible = True
    
End Sub

Private Sub Form_Load()

    'Loads character and picture
    
    C = 1
    
    Open App.Path & "\Characters.txt" For Input As #1
        
        Do Until EOF(1)
            Input #1, Character(C), Life(C), CharacterPic(C)
            C = C + 1
        Loop
    
    Close 1

picCharacter.Picture = LoadPicture(App.Path & "\" & CharacterPic(CharacterNumber))

    
End Sub
