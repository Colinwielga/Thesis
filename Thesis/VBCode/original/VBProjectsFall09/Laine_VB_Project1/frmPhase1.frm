VERSION 5.00
Begin VB.Form frmPhase1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pole Vaulting pg.1"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPhase1.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      Height          =   8655
      Left            =   6000
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H0000C000&
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "Centaur"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton cmdPhase2 
      BackColor       =   &H0000FF00&
      Caption         =   "To Phase 2"
      Height          =   1095
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "    By    Aaron Laine"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8280
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "P     O    L     E            V     A    U   L     T    I     N     G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmPhase1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pole Vaulting'
'Phase 1
'Aaron Laine
'10/19/09
'This Phase will allow the reader to put input into the comuter and it will also send you to phase 2.
Option Explicit
'Hides phase1 and shows phase 2
Private Sub cmdPhase2_Click()
FrmPhase2.Show
frmPhase1.Hide
End Sub
' Allows the reader to input name and town so that a message box pops up'
Private Sub cmdBegin_Click()
Dim Name As String, Town As String

Name = InputBox("Please enter your Name", (Name))
Town = InputBox("Please enter your town", Name)
MsgBox ("Hello, " & Name & " From  " & Town & ".  You are going to learn about Pole Vaulting!")
End Sub


