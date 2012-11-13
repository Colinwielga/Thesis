VERSION 5.00
Begin VB.Form frmopening 
   Caption         =   "Opeing Frame"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   Picture         =   "frmopening.frx":0000
   ScaleHeight     =   10470
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgotomain 
      Caption         =   "Go to Main Screen"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5760
      TabIndex        =   0
      Top             =   7440
      Width           =   3975
   End
   Begin VB.Label lblintro 
      Caption         =   $"frmopening.frx":AD81
      Height          =   735
      Left            =   5520
      TabIndex        =   1
      Top             =   9120
      Width           =   4455
   End
End
Attribute VB_Name = "frmopening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this program will give you a look at the UFC and what its all about

'this button will bring up the main screen from which the project is centered aroung
'it will also hide the this frame
Private Sub cmdgotomain_Click()
frmmainscreen.Show
frmopening.Hide
End Sub
