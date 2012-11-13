VERSION 5.00
Begin VB.Form frmPractice 
   BackColor       =   &H0000C000&
   Caption         =   "Practice Essentials"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackNov 
      Caption         =   "Back to November schedule"
      Height          =   1455
      Left            =   6240
      TabIndex        =   1
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lblPractice 
      Caption         =   "You will need a jersey, shorts, and athletic shoes as well as a completed physical form and parent permission slip. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5880
      TabIndex        =   0
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   0
      Picture         =   "frmPractice.frx":0000
      Top             =   120
      Width           =   5625
   End
End
Attribute VB_Name = "frmPractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBackNov_Click()
frmPractice.Hide
frmNovember.Show
End Sub
