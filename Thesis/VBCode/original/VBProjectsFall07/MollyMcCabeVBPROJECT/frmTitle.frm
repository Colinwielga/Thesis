VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H00400000&
   Caption         =   "Exit Territory"
   ClientHeight    =   6975
   ClientLeft      =   4110
   ClientTop       =   1890
   ClientWidth     =   6075
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   6075
   Begin VB.TextBox txtby 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro M"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4080
      TabIndex        =   4
      Text            =   "By"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000C0&
      Caption         =   "Leave Twins Territory"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H000000C0&
      Caption         =   "Enter Twins Territory"
      BeginProperty Font 
         Name            =   "Kozuka Gothic Pro EL"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5700
      Left            =   960
      Picture         =   "frmTitle.frx":0000
      ScaleHeight     =   2498.339
      ScaleMode       =   0  'User
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   240
      Width           =   4275
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Kozuka Gothic Pro M"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Text            =   "Molly McCabe"
         Top             =   5400
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEnter_Click()
    frmTitle.Hide 'hides Main form
    frmMain.Show 'shows Title form
End Sub

Private Sub cmdQuit_Click()
    End 'ends program
End Sub
