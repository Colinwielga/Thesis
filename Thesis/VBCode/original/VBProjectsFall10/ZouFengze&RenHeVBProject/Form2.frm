VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Witch's Hut"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10710
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   10710
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton btnEnter 
      Caption         =   "Enter"
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "    The door opened automatically. Looking inside, you saw nothing except darkness..."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   7680
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "    In the middle of nowhere, you saw a castle. It's very late; you and your friends decide to spend your night in the hut."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7320
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"Form2.frx":21A46
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6960
      Width           =   9015
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   0
      Picture         =   "Form2.frx":21AF8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnEnter_Click()

    EnterOrNot.Show
End Sub

Private Sub btnMenuClose_Click()
    FrameMenu.Visible = False
    btnSave.Visible = False
    btnLoad.Visible = False
    btnMainMenu.Visible = False
    btnExit.Visible = False
    btnMenuClose.Visible = False
    btnMenuOpen.Visible = True

End Sub

Private Sub btnMenuOpen_Click()

    FrameMenu.Visible = True
    btnSave.Visible = True
    btnLoad.Visible = True
    btnMainMenu.Visible = True
    btnExit.Visible = True
    btnMenuOpen.Visible = False
    btnMenuClose.Visible = True


End Sub

