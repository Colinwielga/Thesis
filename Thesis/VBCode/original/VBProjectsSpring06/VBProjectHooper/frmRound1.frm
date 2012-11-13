VERSION 5.00
Begin VB.Form frmRound1 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWarRoom 
      Caption         =   "To War Room"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   6720
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   4935
      Left            =   7800
      ScaleHeight     =   4875
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   5415
      Left            =   0
      Picture         =   "frmRound1.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   7515
      TabIndex        =   2
      Top             =   0
      Width           =   7575
   End
   Begin VB.CommandButton cmdDraftOrder 
      Caption         =   "See Draft Order"
      Height          =   735
      Left            =   6360
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label lblQuickFacts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Facts"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   -240
      TabIndex        =   0
      Top             =   5640
      Width           =   7575
   End
End
Attribute VB_Name = "frmRound1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdDraftOrder_Click()
    frmRound1.Hide
    frmDraftOrder.Show
End Sub
Private Sub cmdWarRoom_Click()
    frmRound1.Hide
    frmWarRoom.Show
End Sub
