VERSION 5.00
Begin VB.Form frmPromotion 
   BackColor       =   &H00808000&
   Caption         =   "PROMOTION!!!"
   ClientHeight    =   12735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   12735
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitchUp 
      Caption         =   "New Job"
      Height          =   1335
      Left            =   9960
      TabIndex        =   3
      Top             =   10920
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   4200
      Picture         =   "frmPromotion.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "You are truly moving up in the world."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   8880
      Width           =   12255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Congratulations, you have received a promotion for your diligent number-crunching!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   11895
   End
End
Attribute VB_Name = "frmPromotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Bright colors and light-hearted picture reward user for progressing this far.
    
Private Sub cmdSwitchUp_Click()
    'Continue on to more work forms
    frmAnalysis.Show
    frmPromotion.Hide
End Sub
