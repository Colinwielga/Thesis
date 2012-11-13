VERSION 5.00
Begin VB.Form frmBradGustafson1 
   Caption         =   "BradGustasfon1"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   Picture         =   "frmBradGustafson1.frx":0000
   ScaleHeight     =   6195
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStats 
      Caption         =   "2005-06 Stats"
      Height          =   855
      Left            =   7320
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      Height          =   495
      Left            =   8640
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   2
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmd2005DraftPicks 
      Caption         =   "2005 Draft Picks"
      Height          =   855
      Left            =   840
      Picture         =   "frmBradGustafson1.frx":1580F
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H8000000D&
      Caption         =   "Dallas Cowboys"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   1455
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmBradGustafson1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd2005DraftPicks_Click() 'This button shows a form that lists the 2005 draft pick for Dallas'
    frm2005DraftPicks.Show
End Sub

Private Sub cmdExit_Click() 'This button Ends my program'
    End
End Sub

Private Sub cmdStats_Click() 'This button shows a form that has other buttons that contain statistics of players'
    frmBradGustafsonStats.Show
End Sub

