VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form12"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNorth 
      Caption         =   "North"
      Height          =   975
      Left            =   1440
      TabIndex        =   4
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdSouth 
      Caption         =   "South"
      Height          =   855
      Left            =   1440
      TabIndex        =   3
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "Talk to the Imp"
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   4440
      ScaleHeight     =   3795
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   $"Form12.frx":0000
      Height          =   1695
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImp_Click()
'This is for a quick victory'
If Treasure >= 800 Then
    MsgBox "The imp takes all of your money and says, this is exactly the right amount."
    Form12.Hide
    Form16.Show
ElseIf Treasure < 800 Then
    MsgBox "The imp laughs and says, 'You don't have enough money.' and Promptly dissapears."
    cmdImp.Visible = False
End If
End Sub

Private Sub cmdNorth_Click()
'movement function'
Form12.Hide
Form13.Show
End Sub

Private Sub cmdSouth_Click()
'To prevent players from wasting too much time on the simple game'
MsgBox "Why go south? There's nothing but broken rocks. Just go north...seriously."
End Sub

Private Sub Form_Load()
'pic load'
picDungeon.Picture = LoadPicture("Dungeon8d.jpg")
End Sub
