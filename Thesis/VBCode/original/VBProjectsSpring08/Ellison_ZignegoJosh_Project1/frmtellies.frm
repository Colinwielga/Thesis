VERSION 5.00
Begin VB.Form frmtellies 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Tellies"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Continue on your tour de st. joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdorder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Can I take your order?"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lbldumb 
      BackColor       =   &H00C0FFC0&
      Height          =   7455
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lelwelcome 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Welcome to Tellies"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   6795
      Left            =   960
      Picture         =   "frmtellies.frx":0000
      Top             =   1440
      Width           =   9060
   End
End
Attribute VB_Name = "frmtellies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
    'Project name:  Tour De St. Joe
    'Form:  frmtellies, "Tellies"
    'Author:  Brooke
    'Date:  3/11/08
    'Objective: To be able to navigate between this form and the main one.

Private Sub cmdhang_Click()

    frmtelliehang.Show
    frmtellies.Hide

End Sub

Private Sub cmdleave_Click()

    frmjoetown.Show
    frmtellies.Hide

End Sub

Private Sub cmdorder_Click()

    frmtellieorder.Show
    frmtellies.Hide

End Sub

