VERSION 5.00
Begin VB.Form frmDungeon5 
   BackColor       =   &H000000C0&
   Caption         =   "Final Room"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDungeon5 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdSubmit5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdLook5 
      BackColor       =   &H00C0C000&
      Caption         =   "Look Around"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit5 
      BackColor       =   &H000000FF&
      Caption         =   "Quit Program"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.PictureBox picDungeon5 
      Height          =   2655
      Left            =   1080
      ScaleHeight     =   2595
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmDungeon5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit5_Click()
End
End Sub
