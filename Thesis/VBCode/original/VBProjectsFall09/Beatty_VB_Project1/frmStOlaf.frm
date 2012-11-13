VERSION 5.00
Begin VB.Form frmStOlaf 
   BackColor       =   &H00008080&
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   Picture         =   "frmStOlaf.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2400
      ScaleHeight     =   2835
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for information"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   2055
   End
End
Attribute VB_Name = "frmStOLaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:StOlaf
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about St. Olaf
Option Explicit

Private Sub cmdClick_Click()

    Dim Info As String
    
    Open App.Path & "\oles.txt" For Input As #7
    
    Do While Not EOF(7)
        Input #7, Info
        picResults.Print Info
    Loop
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmStOLaf.Hide
    frmMIAC.Show
End Sub
