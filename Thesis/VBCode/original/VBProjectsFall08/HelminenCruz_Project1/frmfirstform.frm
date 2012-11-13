VERSION 5.00
Begin VB.Form frmfirstform 
   BackColor       =   &H00FF00FF&
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form2"
   Picture         =   "frmfirstform.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Log In "
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   2175
   End
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   2880
      Picture         =   "frmfirstform.frx":EBED2
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   720
      Picture         =   "frmfirstform.frx":F42C4
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   2880
      Picture         =   "frmfirstform.frx":FC5CE
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   720
      Picture         =   "frmfirstform.frx":1048E0
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblwelcome 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the Pet Shop"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmfirstform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frmsecondform.Hide
Form2.Show
End Sub
