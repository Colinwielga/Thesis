VERSION 5.00
Begin VB.Form frmUST 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   Picture         =   "frmUST.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for information"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label lblUST4 
      BackColor       =   &H00808080&
      Caption         =   "*Colors: Purple and Gray "
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
      Left            =   2640
      TabIndex        =   6
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label lblUST3 
      BackColor       =   &H00808080&
      Caption         =   "*Stadium: O'Shaughnessy Stadium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   5
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label lblUST2 
      BackColor       =   &H00808080&
      Caption         =   "*Oddly enough the nickname is the Tommies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lblUST1 
      BackColor       =   &H00808080&
      Caption         =   "*St. Thomas is located in St. Paul, MN."
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
      Left            =   2640
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
End
Attribute VB_Name = "frmUST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:UST
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about St. Thomas
Option Explicit

Private Sub cmdClick_Click()
MsgBox "If you weren't there on Saturday, The Johnnies Beat the Tommies in Overtime", , "Amazing!!"
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
   frmUST.Hide
   frmMIAC.Show
End Sub

