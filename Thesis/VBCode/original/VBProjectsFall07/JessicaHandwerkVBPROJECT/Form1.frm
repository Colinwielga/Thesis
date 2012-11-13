VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "frmHabitat"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8490
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
      Left            =   7320
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
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
      Left            =   7320
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lbltext2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2415
      Left            =   1680
      TabIndex        =   2
      Top             =   2280
      Width           =   5415
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":019F
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2895
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label lblHabitat 
      BackStyle       =   0  'Transparent
      Caption         =   "What Is Habitat For Humanity?"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   5790
      Left            =   0
      Picture         =   "Form1.frx":02A5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdMenu_Click()
frmForm1.Show
frmHabitat.Hide

End Sub
