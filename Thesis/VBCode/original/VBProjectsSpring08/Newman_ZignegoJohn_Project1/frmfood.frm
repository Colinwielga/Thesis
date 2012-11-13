VERSION 5.00
Begin VB.Form frmfood 
   BackColor       =   &H000080FF&
   Caption         =   "Food"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdg 
      Height          =   3135
      Index           =   0
      Left            =   5400
      Picture         =   "frmfood.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   10575
      Left            =   10560
      ScaleHeight     =   10515
      ScaleWidth      =   7035
      TabIndex        =   8
      Top             =   360
      Width           =   7095
   End
   Begin VB.CommandButton cmdblc 
      Height          =   3495
      Index           =   3
      Left            =   240
      Picture         =   "frmfood.frx":2635
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   4575
   End
   Begin VB.CommandButton cmdpacj 
      DisabledPicture =   "frmfood.frx":E17B
      Height          =   3615
      Index           =   2
      Left            =   5400
      Picture         =   "frmfood.frx":12F29
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdseeds 
      Height          =   3375
      Index           =   1
      Left            =   5520
      Picture         =   "frmfood.frx":17CD7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   2175
   End
   Begin VB.TextBox txtseeds 
      Height          =   495
      Index           =   3
      Left            =   8160
      TabIndex        =   3
      Top             =   9480
      Width           =   2055
   End
   Begin VB.TextBox txtblc 
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox tctpacj 
      Height          =   495
      Index           =   1
      Left            =   8040
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtg 
      Height          =   495
      Index           =   0
      Left            =   8040
      TabIndex        =   0
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label lblfood 
      BackColor       =   &H000080FF&
      Caption         =   "Food"
      BeginProperty Font 
         Name            =   "MS PMincho"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1080
      TabIndex        =   7
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "frmfood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

