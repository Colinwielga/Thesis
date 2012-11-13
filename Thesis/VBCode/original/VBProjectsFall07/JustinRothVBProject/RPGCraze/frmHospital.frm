VERSION 5.00
Begin VB.Form frmHospital 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hospital"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAgility 
      Caption         =   "Increase your Agility!"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdHeal 
      Caption         =   "Heal Yourself"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblLogo2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLogo1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblHospital 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: RPGCraze
'Form name: frmHospital
'Author: Justin Roth
'Date Written: Sunday, November 4th, 2007
'Objective of form: This form is where the user can choose to heal themselves, as well as increase their agility points.
        'The user could come here after leaving the Quest form.

Option Explicit

Private Sub cmdBack_Click()
    frmHospital.Hide    'Goes back to the Map form.
End Sub

Private Sub cmdHeal_Click()
    'If the user's health is less than 100, then they can heal themselves back to full health.
    If MyHealth < 100 Then
        MyHealth = 100
        MsgBox "You have been fully healed!", , "100% Health!"
    Else
        MsgBox "You are already at full health."    'If the user already has 100 health, then they are notified.
    End If
    
End Sub

Private Sub cmdAgility_Click()
    'The user can increase their agility points, but only up to five points.
    If Agility < 5 Then
        Agility = Agility + 1
        MsgBox "You gained 1 agility point!", , "+1 Agility!"
    Else
        MsgBox "You can not increase your agility anymore.", , "Full Agility!"  'If the user has all five agility points, then they are told that they can't increase anymore.
    End If
    
End Sub
