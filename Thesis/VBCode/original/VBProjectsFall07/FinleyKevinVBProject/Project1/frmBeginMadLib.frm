VERSION 5.00
Begin VB.Form frmBeginMadLib 
   Caption         =   "Start Page"
   ClientHeight    =   7515
   ClientLeft      =   3960
   ClientTop       =   2115
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   7500
   Begin VB.PictureBox picAmericanSP 
      Height          =   7695
      Left            =   -360
      Picture         =   "FRMBEG~1.frx":0000
      ScaleHeight     =   7635
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   -120
      Width           =   9975
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdGWB 
         Height          =   2895
         Left            =   600
         Picture         =   "FRMBEG~1.frx":63E15
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3480
         Width           =   3375
      End
      Begin VB.CommandButton cmdBC 
         Height          =   2895
         Left            =   4200
         Picture         =   "FRMBEG~1.frx":7980C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3480
         Width           =   3375
      End
      Begin VB.CommandButton cmdJFK 
         Height          =   2895
         Left            =   600
         Picture         =   "FRMBEG~1.frx":8EA00
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmdFDR 
         Height          =   2895
         Left            =   4200
         Picture         =   "FRMBEG~1.frx":A0E8B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblBegin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click on a president to take a short quiz, and based on your answers you will be able to edit their Inauguration Speech ."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   5
         Top             =   6480
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmBeginMadLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBC_Click()
    frmQuizBC.Show
    frmBeginMadLib.Hide
End Sub
Private Sub cmdFDR_Click()
    frmQuizFDR.Show
    frmBeginMadLib.Hide
End Sub
Private Sub cmdGWB_Click()
    frmQuizGWB.Show
    frmBeginMadLib.Hide
End Sub

Private Sub cmdJFK_Click()
    frmQuizJFK.Show
    frmBeginMadLib.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
