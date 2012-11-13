VERSION 5.00
Begin VB.Form frmMain
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdE
      Caption         =   "Exit"
      Height          =   855
      Left            =   1080
      TabIndex        =   4
      Top             =   7080
      Width           =   4095
   End
   Begin VB.CommandButton cmdww
      Caption         =   "Weapon Calculator"
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdG
      Caption         =   "Power and Weapon(Quiz)"
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdStore
      Caption         =   "Kick-Ass Store"
      Height          =   975
      Left            =   3480
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdC
      Caption         =   "Characters Overview"
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Image Image1
      Height          =   9000
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Top             =   0
      Width           =   6075
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Sub Command1_Click()

    End Sub

    Private Sub cmdC_Click()
        frmMain.Hide
        frmCharacters.Show
    End Sub

    Private Sub cmdE_Click()
    End
    End Sub

    Private Sub Command2_Click()

    End Sub

    Private Sub cmdG_Click()
    frmMain.Hide
    frmData.Show
    End Sub

    Private Sub cmdStore_Click()
    frmMain.Hide
    frmStore.Show
    End Sub

    Private Sub cmdww_Click()
    frmMain.Hide
    frmCal.Show
    End Sub
