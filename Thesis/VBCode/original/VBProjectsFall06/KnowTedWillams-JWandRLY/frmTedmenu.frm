VERSION 5.00
Begin VB.Form frmTedmenu 
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   Picture         =   "frmTedmenu.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcareerfrm 
      BackColor       =   &H000000FF&
      Caption         =   "Ted's Highlights"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   4095
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H000000FF&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   4095
   End
   Begin VB.CommandButton cmdstorefrm 
      BackColor       =   &H000000FF&
      Caption         =   "Ted's Store"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   4095
   End
   Begin VB.CommandButton cmdstatfrm 
      BackColor       =   &H000000FF&
      Caption         =   "Ted's Stats"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton cmdpicsfrm 
      BackColor       =   &H000000FF&
      Caption         =   "Ted's Pics"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   4095
   End
   Begin VB.CommandButton cmdbiofrm 
      BackColor       =   &H000000FF&
      Caption         =   "Ted's Bio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "frmTedmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbiofrm_Click()
    frmTedBio.Show
    frmTedmenu.Hide
End Sub


Private Sub cmdcareerfrm_Click()
    frmTedCareerHighs.Show
    frmTedmenu.Hide
End Sub

Private Sub cmdpicsfrm_Click()
    frmTedPics.Show
    frmTedmenu.Hide
End Sub

Private Sub cmdquit_Click()
    End
End Sub

Private Sub cmdstatfrm_Click()
    frmTedStats.Show
    frmTedmenu.Hide
End Sub

Private Sub cmdstorefrm_Click()
    frmTedStore.Show
    frmTedmenu.Hide
End Sub
