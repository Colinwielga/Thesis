VERSION 5.00
Begin VB.Form Maybach 
   BackColor       =   &H80000008&
   Caption         =   "Maybach 57"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   10560
      TabIndex        =   1
      Top             =   7800
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   7335
      Left            =   120
      Picture         =   "frm34.frx":0000
      ScaleHeight     =   7275
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   120
      Width           =   11415
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Maybach 57"
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
      Top             =   7560
      Width           =   2295
   End
End
Attribute VB_Name = "Maybach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Maybach.Hide
End Sub
