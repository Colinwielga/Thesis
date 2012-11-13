VERSION 5.00
Begin VB.Form Ram 
   BackColor       =   &H80000008&
   Caption         =   "Dodge Ram"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   10200
      TabIndex        =   1
      Top             =   6000
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   120
      Picture         =   "frm16.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   10995
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Dodge Ram"
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
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   2295
   End
End
Attribute VB_Name = "Ram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Ram.Hide
End Sub

