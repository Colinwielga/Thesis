VERSION 5.00
Begin VB.Form A8 
   BackColor       =   &H80000008&
   Caption         =   "Audi A8"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   240
      Picture         =   "frm30.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Audi A8"
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
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "A8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    A8.Hide
End Sub
