VERSION 5.00
Begin VB.Form frmReebok1 
   Caption         =   "Reebok1"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   Picture         =   "frmReebok1.frx":0000
   ScaleHeight     =   12000
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdComputeAge 
      BackColor       =   &H000000FF&
      Caption         =   "Compute Age"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdBackHome 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Go To Store Home"
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FF00&
      Caption         =   "Quit"
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10080
      Width           =   1815
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label lblReebok 
      Caption         =   "Reebok"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblAge 
      BackColor       =   &H008080FF&
      Caption         =   "Enter your Age:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "frmReebok1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' AthleticStore
' Reebok1
' Andrew Eisinger
' 3/17/09
'This program gets the input via a text box
'This program then based on the input sends to a different form

Dim Age As Single

Private Sub cmdBackHome_Click()
frmStoreHome.Show
frmReebok1.Hide
End Sub

Private Sub cmdComputeAge_Click()
Age = txtAge.Text
If Age >= 18 And Age <= 39 Then
    frmReebokAdults.Show
    frmReebok1.Hide
ElseIf Age <= 17 Then
    frmReebokKids.Show
    frmReebok1.Hide
ElseIf Age >= 40 Then
    frmReebokSeniors.Show
    frmReebok1.Hide
End If

End Sub

Private Sub cmdQuit_Click()
End
End Sub
