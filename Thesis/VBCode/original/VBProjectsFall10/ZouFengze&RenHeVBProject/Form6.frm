VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Witch's hut"
   ClientHeight    =   7980
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10710
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10710
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton btnNo 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btnYes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "    Some noise wakes you up. Those noise sounds from a box in the coner. Do you want to open the wooden box?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   3135
      Left            =   8520
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   $"Form6.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   8055
   End
   Begin VB.Image Image1 
      Height          =   6300
      Left            =   0
      Picture         =   "Form6.frx":00C4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8205
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNext_Click()
    Label2.Visible = True
    btnYes.Visible = True
    btnNo.Visible = True
    btnNext.Enabled = False
End Sub

Private Sub btnNo_Click()
    MsgBox "You cannot going to sleep again, because of the noise."
End Sub

Private Sub btnYes_Click()
    MsgBox "When you open the wooden box, you can hear a gilr crying!"
    MsgBox "After searching the box, nothing in it, but a book."
    Form6.Hide
    Form7.Show
    
End Sub
