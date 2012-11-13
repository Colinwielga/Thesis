VERSION 5.00
Begin VB.Form frmDungeon6 
   BackColor       =   &H00000000&
   Caption         =   "The Final Room"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDungeon6 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Text            =   " "
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdSubmit6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdLook6 
      BackColor       =   &H00C0C000&
      Caption         =   "Look Around"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnd6 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit Dungeon"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit6 
      BackColor       =   &H000000FF&
      Caption         =   "Quit Program"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdLoadDungeon6 
      Caption         =   "Load"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox picDungeon6 
      Height          =   2655
      Left            =   1080
      ScaleHeight     =   2595
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmDungeon6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
