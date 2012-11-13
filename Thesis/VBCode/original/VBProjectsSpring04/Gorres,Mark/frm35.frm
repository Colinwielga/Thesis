VERSION 5.00
Begin VB.Form Enzo 
   BackColor       =   &H80000008&
   Caption         =   "Ferrari Enzo"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   10200
      TabIndex        =   1
      Top             =   7200
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   6735
      Left            =   120
      Picture         =   "frm35.frx":0000
      ScaleHeight     =   6675
      ScaleWidth      =   10995
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Ferrari Enzo"
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
      Top             =   6960
      Width           =   2535
   End
End
Attribute VB_Name = "Enzo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Enzo.Hide
End Sub
