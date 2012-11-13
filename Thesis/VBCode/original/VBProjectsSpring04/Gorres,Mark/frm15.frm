VERSION 5.00
Begin VB.Form Liberty 
   BackColor       =   &H80000008&
   Caption         =   "Jeep Liberty"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   4320
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   240
      Picture         =   "frm15.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Jeep Liberty"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Width           =   2655
   End
End
Attribute VB_Name = "Liberty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Liberty.Hide
End Sub

