VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H00400000&
   Caption         =   "Twins Territory"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
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
      Picture         =   "frmM.frx":0000
      ScaleHeight     =   2498.339
      ScaleMode       =   0  'User
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   240
      Width           =   4275
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
    MsgBox "The Best Starting Lineup for the Twins in 2007. (in my opinion at least)", , "Starting Lineup"
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
