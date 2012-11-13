VERSION 5.00
Begin VB.Form Focus 
   BackColor       =   &H80000008&
   Caption         =   "Ford Focus"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   6360
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   120
      Picture         =   "frm3.frx":0000
      ScaleHeight     =   5835
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Ford Focus"
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
      Top             =   6120
      Width           =   2175
   End
End
Attribute VB_Name = "Focus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Focus.Hide
End Sub
