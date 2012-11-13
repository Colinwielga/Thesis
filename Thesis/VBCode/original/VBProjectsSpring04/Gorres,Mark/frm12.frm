VERSION 5.00
Begin VB.Form Accord 
   BackColor       =   &H80000008&
   Caption         =   "Honda Accord"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   6000
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   120
      Picture         =   "frm12.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Honda Accord"
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
      Top             =   5880
      Width           =   2775
   End
End
Attribute VB_Name = "Accord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Accord.Hide
End Sub

