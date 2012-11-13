VERSION 5.00
Begin VB.Form frmENTER 
   BackColor       =   &H8000000D&
   Caption         =   "WORLD OF POKEMON ENTRANCE"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   FillColor       =   &H00C00000&
   ForeColor       =   &H80000002&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H80000011&
      Caption         =   "CLICK HERE to enter the exciting World of POKEMON!"
      Height          =   2175
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "THE WORLD OF POKEMON: A SIMULATION"
      BeginProperty Font 
         Name            =   "OCR A Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblIntro 
      BackColor       =   &H8000000D&
      Caption         =   "Project by Peter Smorynski. Based on the ""Pokemon"" franchise copyrighted by NINTENDO"
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   8640
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   480
      Picture         =   "Form1bigone.frx":0000
      Top             =   480
      Width           =   9360
   End
End
Attribute VB_Name = "frmENTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Text1_Change()

End Sub

Private Sub cmdEnter_Click()
Dim cmdEnter
frmENTER.Hide
frmCHARACTER.Show
MsgBox ("Let's begin the simulation by having you select your Pokemon Trainer and giving them a name!"), , ("INSTRUCTION!")
End Sub
