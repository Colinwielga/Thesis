VERSION 5.00
Begin VB.Form frmMacalaster 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   Picture         =   "frmMacalaster.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtM3 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Text            =   "Colors: Blue and Orange"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox txtM2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Text            =   "Nickname: Scots"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtM1 
      BackColor       =   &H000080FF&
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
      Left            =   2760
      TabIndex        =   3
      Text            =   "Location:St. Paul, MN"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for more information"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   2175
   End
End
Attribute VB_Name = "frmMacalaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClick_Click()
MsgBox "The Scots play at Macalaster Stadium", , "Crazy stuff"
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmMacalaster.Hide
    frmMIAC.Show
End Sub

