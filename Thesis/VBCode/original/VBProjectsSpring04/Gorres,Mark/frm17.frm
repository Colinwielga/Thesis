VERSION 5.00
Begin VB.Form F150 
   BackColor       =   &H80000008&
   Caption         =   "Ford F-150"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   8640
      TabIndex        =   1
      Top             =   7080
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Left            =   120
      Picture         =   "frm17.frx":0000
      ScaleHeight     =   6555
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Ford F-150"
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
      Top             =   6840
      Width           =   2055
   End
End
Attribute VB_Name = "F150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    F150.Hide
End Sub

