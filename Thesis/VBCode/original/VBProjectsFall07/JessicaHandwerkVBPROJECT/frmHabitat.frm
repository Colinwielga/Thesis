VERSION 5.00
Begin VB.Form frmHabitat 
   Caption         =   "Habitat for Humanity"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblText3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "At Central Minnesoat Habitat for Humanity, Building Homes and Building Hope is our only focus!"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   5400
      Width           =   6855
   End
   Begin VB.Label lbltext2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHabitat.frx":0000
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   6480
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblText1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHabitat.frx":016C
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblHabitat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What is Habitat for Humanity?"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   8760
      Left            =   0
      Picture         =   "frmHabitat.frx":0274
      Top             =   -720
      Width           =   8625
   End
End
Attribute VB_Name = "frmHabitat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdMenu_Click()
frmMenu.Show
frmHabitat.Hide
End Sub
