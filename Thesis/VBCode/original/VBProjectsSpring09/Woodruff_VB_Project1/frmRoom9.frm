VERSION 5.00
Begin VB.Form frmRoom9 
   BackColor       =   &H80000017&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtD 
      Height          =   600
      Left            =   11880
      TabIndex        =   7
      Top             =   4800
      Width           =   2500
   End
   Begin VB.TextBox txtC 
      Height          =   600
      Left            =   11880
      TabIndex        =   6
      Top             =   4080
      Width           =   2500
   End
   Begin VB.TextBox txtB 
      Height          =   600
      Left            =   11880
      TabIndex        =   5
      Top             =   3360
      Width           =   2500
   End
   Begin VB.TextBox txtA 
      Height          =   600
      Left            =   11880
      TabIndex        =   4
      Top             =   2640
      Width           =   2500
   End
   Begin VB.CommandButton cmdWin 
      BackColor       =   &H80000015&
      Caption         =   "Tell her the words."
      Height          =   800
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   2500
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000015&
      Caption         =   "Go Back"
      Height          =   800
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2500
   End
   Begin VB.PictureBox picRoom9 
      Height          =   5175
      Left            =   3120
      ScaleHeight     =   5115
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   960
      Width           =   8535
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
      Left            =   11880
      TabIndex        =   9
      Top             =   1800
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
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblStoryRoom9 
      BackColor       =   &H80000017&
      Caption         =   $"frmRoom9.frx":0000
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
      Left            =   2760
      TabIndex        =   3
      Top             =   6600
      Width           =   8175
   End
End
Attribute VB_Name = "frmRoom9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom9
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is a 'room' of the game.  There is a witch here that gives the user a
'secret spell.

Option Explicit
Dim Code1 As String, Code2 As String, Code3 As String, Code4 As String
Dim CodeA As String, CodeB As String, CodeC As String, CodeD As String
Dim Random As String

Private Sub cmdBack_Click()

    'Moves user to room room 2 and reminds him/her of the secret code
    frmRoom9.Visible = False
    frmRoom2.Visible = True
    
    If Secret = True Then
        MsgBox "Witch: Don't forget: " & SecretCode & "." & " It might be important", , ""
    End If
    
End Sub


Private Sub cmdWin_Click()

    'The witch gives the user a secret spell based on the words he/she gives her
    'It's used for the dragon at the end
    Code1 = txtA.Text
    Code2 = txtB.Text
    Code3 = txtC.Text
    Code4 = txtD.Text
    
    CodeA = Left(Code1, 3)
    CodeB = Right(Code2, 3)
    CodeC = Mid(Code3, 2)
    CodeD = Left(Code4, 1)
    
    Randomize
    Random = Int(Rnd * 10)
    
    SecretCode = CodeA & CodeB & CodeC & CodeD & Random
    MsgBox "Witch: Ok, well the magic spell I cooked up for you is " & SecretCode & "." & " Have fun with it.", , ""
    
    Secret = True
    
End Sub

Private Sub Form_Load()

    picRoom9.Picture = LoadPicture(App.Path & "\witch.jpg")
    
End Sub

