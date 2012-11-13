VERSION 5.00
Begin VB.Form A4 
   BackColor       =   &H80000008&
   Caption         =   "Audi A4"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   3240
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   240
      Picture         =   "frm22.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Audi A4"
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
Attribute VB_Name = "A4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    A4.Hide
End Sub

