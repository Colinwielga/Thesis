VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   2430
   ClientTop       =   2595
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9690
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H80000009&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "MaestroTimes"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2160
      TabIndex        =   0
      Top             =   1800
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdContinue_Click()
frmGenres.Show
frmAbout.Hide
End Sub
