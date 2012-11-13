VERSION 5.00
Begin VB.Form frmEquipment 
   BackColor       =   &H0080FF80&
   Caption         =   "Equipment"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   ForeColor       =   &H0080FF80&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7305
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   6480
      ScaleHeight     =   5835
      ScaleWidth      =   3795
      TabIndex        =   9
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   3480
      TabIndex        =   8
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3480
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdCleats 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cleats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdBalls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Balls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton cmdGloves 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gloves"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtEquipment 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Text            =   "Equipment"
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdBats 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
End
Attribute VB_Name = "frmEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
