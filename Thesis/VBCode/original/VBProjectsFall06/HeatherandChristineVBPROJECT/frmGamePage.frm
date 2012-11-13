VERSION 5.00
Begin VB.Form frmGamePage 
   BackColor       =   &H00C00000&
   Caption         =   "Form2"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13815
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcelebform 
      BackColor       =   &H00C000C0&
      Caption         =   "SEE PAST CELEBRITIES!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5400
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8520
      Width           =   5055
   End
   Begin VB.CommandButton cmdEndJeopardy 
      Caption         =   "End"
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   11040
      Width           =   3855
   End
   Begin VB.CommandButton cmdLetsPlay 
      BackColor       =   &H00FFFF00&
      Caption         =   "Lets Play JEOPARDY!"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   8775
   End
   Begin VB.ComboBox Gender 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5760
      TabIndex        =   2
      Text            =   "Chose Gender"
      ToolTipText     =   "Chose a Gender"
      Top             =   4200
      Width           =   3855
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter Name"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6120
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.Frame frmJeopardy 
      BackColor       =   &H00C00000&
      Caption         =   "CSB/SJU JEOPARDY"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   72
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   2175
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   12015
   End
End
Attribute VB_Name = "frmGamePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Game Page:  This form is the first form to display to the user. This is where
'            the user will enter their name and select a gender. The user has
'            the choice of starting the game or looking at what celebrities
'            were on the show.

Private Sub cmdcelebform_Click()
'This button will bring the player to the celebrity form page so he or she can
'   view when certain celebrities were on the show. The first form will disappear
'   and it will automatically bring the user to the celebrity page.

    frmGamePage.Hide
    frmCelebs.Show
End Sub

Private Sub cmdLetsPlay_Click()
'The Lets Play button will bring the player to the game board so the player may begin
'   to play the game. If the player comes back to this form, they will be able to
'   start another game by selecting this button.

    frmGamePage.Hide
    frmJeopardyBoard.Show
    frmJeopardyBoard.txtName.Text = PlayerName
    If Gender.Text = "Female" Then
        frmJeopardyBoard.picGender.Picture = LoadPicture(App.Path & "\girl.jpg")
    Else
        frmJeopardyBoard.picGender.Picture = LoadPicture(App.Path & "\boy.gif")
    End If
    
End Sub


Private Sub cmdName_Click()
' The first thing the player should do is enter their name and this is where the
'   player should do that. The name the player types in will show up on the game
'   form and also on the check when the player is ready to finish the game.

    PlayerName = InputBox("Please enter your name.", "Name")
    frmJeopardyBoard.Caption = "Welcome" & " " & PlayerName
    
End Sub


Private Sub Form_Load()
    Gender.AddItem ("Female")
    Gender.AddItem ("Male")
    
End Sub

