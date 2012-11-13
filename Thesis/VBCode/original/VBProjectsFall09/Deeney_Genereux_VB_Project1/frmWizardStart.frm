VERSION 5.00
Begin VB.Form frmWizardStart 
   BackColor       =   &H00000000&
   Caption         =   "Wizard!"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next!"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWizardStart.frx":0000
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   5310
      Left            =   4800
      Picture         =   "frmWizardStart.frx":00A2
      Top             =   480
      Width           =   3435
   End
End
Attribute VB_Name = "frmWizardStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'In this form, it tells the user what is coming up in other forms

Private Sub CmdNext_Click()
    frmWizardStart.Hide
    frmMountains.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
