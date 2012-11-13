VERSION 5.00
Begin VB.Form frmLanguageMenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Language Menu"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   DrawStyle       =   6  'Inside Solid
   FillStyle       =   5  'Downward Diagonal
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00008000&
      Caption         =   "Help"
      Height          =   855
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnglish 
      BackColor       =   &H00C00000&
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      MaskColor       =   &H00004000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblEnglish 
      BackColor       =   &H80000009&
      Caption         =   "Proceed to the English Menu:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblHelp 
      BackColor       =   &H80000009&
      Caption         =   "Proceed to the Help Dialog:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblMyProgram 
      BackColor       =   &H80000009&
      Caption         =   "VB Alarm Clock - Madeleine Ebacher"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmLanguageMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB Alarm Clock (project1.vbp)
'"Language Menu" (frmLanuageMenu.frm)
'designed by: Madeleine Ebacher
'3/24/06
'This menu allows either immediate access to the help or English menus, or for the user to Exit.
'This project was designed for my COMSCI 130 VB Project, and also because I'm sick of finding such crappy PC Alarm Clocks online! I asked myself, how hard can it be to design an alarm clock? The answer?
'VERY hard.

Private Sub cmdEnglish_Click()
    frmSetClock.Show
    frmLanguageMenu.Hide
End Sub

Private Sub cmdExit_Click()
B = MsgBox("Are You Sure You Want To Exit?! This is a really cool program!", vbQuestion + vbYesNo, "Exit ?")
If B = 6 Then
    End
  Else
 frmLanguageMenu.Show
  End If
End Sub

Private Sub cmdHelp_Click()
    frmLanguageMenu.Hide
    frmHelp.Show
End Sub

