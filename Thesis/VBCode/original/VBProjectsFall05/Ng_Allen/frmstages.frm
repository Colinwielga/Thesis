VERSION 5.00
Begin VB.Form frmstages 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Choose your Stage!"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   7080
      ScaleHeight     =   3795
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "Quit Game :("
      Height          =   2000
      Left            =   5160
      TabIndex        =   2
      Top             =   4320
      Width           =   4000
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Go Back"
      Height          =   2000
      Left            =   480
      TabIndex        =   1
      Top             =   4320
      Width           =   4000
   End
   Begin VB.CommandButton cmdeasy 
      BackColor       =   &H0000C000&
      Caption         =   "Race Type - ER"
      Height          =   1000
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   3000
   End
End
Attribute VB_Name = "frmstages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
    frmstages.Visible = False
    frmMainmenu.Visible = True
End Sub

Private Sub cmdeasy_Click()
    frmstages.Visible = False
    frm1player.Visible = True
    
End Sub
