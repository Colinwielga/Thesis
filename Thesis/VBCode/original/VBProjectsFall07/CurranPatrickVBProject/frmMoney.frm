VERSION 5.00
Begin VB.Form frmMoney 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   Picture         =   "frmMoney.frx":0000
   ScaleHeight     =   5340
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd200 
      Caption         =   "$200"
      Height          =   855
      Left            =   3960
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmd175 
      Caption         =   "$175"
      Height          =   855
      Left            =   1320
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmd150 
      Caption         =   "$150"
      Height          =   855
      Left            =   4800
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmd125 
      Caption         =   "$125"
      Height          =   855
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmd100 
      Caption         =   "$100"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmd75 
      Caption         =   "$75"
      Height          =   975
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmd50 
      Caption         =   "$50"
      Height          =   975
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmd25 
      BackColor       =   &H00000000&
      Caption         =   "$25"
      Height          =   975
      Left            =   360
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H80000012&
      Caption         =   "How much money would you like to put into the machine?"
      BeginProperty Font 
         Name            =   "Blackoak Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   1920
      TabIndex        =   8
      Top             =   4200
      Width           =   4215
   End
End
Attribute VB_Name = "frmMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd100_Click()  'This form asks the player how much money he/she would like to start with.
                            'It then takes whatever button the user presses and gives the player the corresponding amount of credits.
    
Credits = 100
frmMoney.Hide
frmGame.Show
End Sub

Private Sub cmd125_Click()
Credits = 125
frmMoney.Hide
frmGame.Show
End Sub

Private Sub cmd150_Click()
Credits = 150
frmMoney.Hide
frmGame.Show
End Sub

Private Sub cmd175_Click()
Credits = 175
frmMoney.Hide
frmGame.Show
End Sub

Private Sub cmd200_Click()
Credits = 200
frmMoney.Hide
frmGame.Show
End Sub

Private Sub cmd25_Click()
Credits = 25
frmMoney.Hide
frmGame.Show
End Sub

Private Sub cmd50_Click()
Credits = 50
frmMoney.Hide
frmGame.Show
End Sub

Private Sub cmd75_Click()
Credits = 75
frmMoney.Hide
frmGame.Show
End Sub
