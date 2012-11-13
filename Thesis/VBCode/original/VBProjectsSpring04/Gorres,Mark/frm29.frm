VERSION 5.00
Begin VB.Form Navigator 
   BackColor       =   &H80000008&
   Caption         =   "Lincoln Navigator"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   5880
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   120
      Picture         =   "frm29.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Lincoln Navigator"
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
      Top             =   5640
      Width           =   3495
   End
End
Attribute VB_Name = "Navigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Navigator.Hide
End Sub
