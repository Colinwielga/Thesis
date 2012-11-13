VERSION 5.00
Begin VB.Form frmBethel 
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   Picture         =   "frmBethel.frx":0000
   ScaleHeight     =   4815
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtB3 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Text            =   "Staium: Royal Stadium"
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox txtB2 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Text            =   "Located in St. Paul, MN"
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox txtB1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Text            =   "Colors:Royal Blue and Gold"
      Top             =   2400
      Width           =   3495
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for more information"
      Height          =   1095
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "frmBethel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Project name: quick Facts about the MIAC'
    'Form name:Bethel
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about Bethel
Option Explicit


Private Sub cmdClick_Click()
MsgBox "Bethel's nickname is the Royals", , "Guess What!" 'Create a message box for output'
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmMIAC.Show
    frmBethel.Hide
    
End Sub

