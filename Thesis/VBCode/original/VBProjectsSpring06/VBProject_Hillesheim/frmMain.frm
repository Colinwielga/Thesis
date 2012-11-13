VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000001&
   Caption         =   "Choose an Activity"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form2"
   ScaleHeight     =   7800
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWelcome 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   480
      ScaleHeight     =   360
      ScaleWidth      =   8175
      TabIndex        =   4
      Top             =   120
      Width           =   8175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   7560
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmdStats 
      Caption         =   "Get Battle Statistics"
      Height          =   975
      Left            =   7560
      TabIndex        =   1
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Learn About Famous Commanders"
      Height          =   975
      Left            =   7560
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   6975
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "What Do You Want To Do?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7200
      TabIndex        =   5
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000001&
      Caption         =   "By Jacob Hillesheim"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   7200
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Naval History (Naval.vpb)
'Main Page (frmMain.frm)
'Jacob Hillesheim
'March 20,2006
'The purpose of this project is to engage and educate the user about World War II in the Pacific.
'Through this project, users can identify and learn about famous naval leaders of the war
'and identify and learn about specific battles, their results, and their ramifications on the war as a whole
'This form is to welcome the user and has command buttons which lead the user to different activities

Private Sub cmdInfo_Click()
    'Takes user to Famous Leader page
    frmMain.Hide
    frmLeader.Show
End Sub
Private Sub cmdQuit_Click()
    'Ends program
    End
End Sub
Private Sub cmdStats_Click()
    'Takes user to Battle page
    frmMain.Hide
    frmBattle.Show
End Sub

