VERSION 5.00
Begin VB.Form frmStartup 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Startup"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   5640
      Picture         =   "frmStartup.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search for your Favorite Player"
      Height          =   975
      Index           =   1
      Left            =   7800
      TabIndex        =   4
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdNFL 
      BackColor       =   &H00FFFFC0&
      Caption         =   "NFL"
      Height          =   975
      Index           =   7
      Left            =   5640
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdAFC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "AFC"
      Height          =   975
      Index           =   6
      Left            =   9600
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      Height          =   975
      Index           =   5
      Left            =   3360
      TabIndex        =   1
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdNFC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "NFC"
      Height          =   975
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project-NFL Stats
'Form-frmStartup
'Written by Ryan Kempf and Ryan Dell
'3-22-09
'This form is our main menu and allows a user to access different options they deem necessary

Private Sub cmdAFC_Click(Index As Integer)
    frmStartup.Hide
    frmAFC.Show
End Sub


Private Sub cmdNFC_Click(Index As Integer)
    frmStartup.Hide
    frmNFC.Show
End Sub


Private Sub cmdNFL_Click(Index As Integer)
    frmStartup.Hide
    frmNFL.Show
End Sub

Private Sub cmdQuit_Click(Index As Integer)
    End
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    frmStartup.Hide
    frmSearch.Show
End Sub
