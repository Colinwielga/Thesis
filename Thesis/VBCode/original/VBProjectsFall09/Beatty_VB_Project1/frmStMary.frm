VERSION 5.00
Begin VB.Form frmStCates 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   Picture         =   "frmStMary.frx":0000
   ScaleHeight     =   6075
   ScaleWidth      =   7005
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
      Height          =   1815
      Left            =   2160
      ScaleHeight     =   1755
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for information"
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label lblcath1 
      BackColor       =   &H00800080&
      Caption         =   "Though St. Catherine has men and women students, it only has womens sports. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2160
      TabIndex        =   4
      Top             =   4680
      Width           =   2655
   End
End
Attribute VB_Name = "frmStCates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:StCates
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about St. Catherine'
Option Explicit

Private Sub cmdClick_Click()

    Dim Info As String
    
    Open App.Path & "\cath.txt" For Input As #5
    
    
    Do While Not EOF(5)
        Input #5, Info
        picResults.Print Info
    Loop
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmStMary.Hide
    frmMIAC.Show
End Sub
