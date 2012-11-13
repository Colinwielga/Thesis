VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00008000&
   Caption         =   "Help Dialog"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpenHelp 
      BackColor       =   &H0000FFFF&
      Caption         =   "CLICK ME!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdExit3 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to English Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H000080FF&
      Caption         =   "Search Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox txtHelping 
      Height          =   4335
      HideSelection   =   0   'False
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmHelp.frx":0000
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label lblMyProgram 
      BackColor       =   &H00008000&
      Caption         =   "VB Alarm Clock - Madeleine Ebacher"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   6120
      Width           =   2655
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB Alarm Clock (project1.vbp)
'"Help Menu" (frmHelp.frm)
'designed by: Madeleine Ebacher
'3/24/06
'This menu allows the user to search the Help document for keywords. Or at least it would if I could get it to work!

Option Explicit
Dim B As Integer
Dim Counter As Integer
Dim Found As Boolean
Dim N As String
Dim Help(1 To 85) As String
Dim I As Integer

Private Sub cmdExit3_Click()
B = MsgBox("Are You Sure You Want To Exit?! This is a really cool program!", vbQuestion + vbYesNo, "Exit ?")
If B = 6 Then
    End
  Else
 frmLanguageMenu.Show
  End If
End Sub

Private Sub cmdOpenHelp_Click()
cmdOpenHelp.Visible = False
End Sub

Private Sub cmdReturn_Click()
    frmHelp.Hide
    frmEnglishMenu.Show
End Sub

Private Sub cmdSearch_Click()
 Open App.Path & "\VBClockHelp.txt" For Input As #1
    N = txtSearch.Text
    If N = "" Then
        MsgBox "please enter text to search", "Oops!"
    End If
    
    Found = False
    I = 0

Do While ((Not Found) And (I < 86))
    I = I + 1
    If N = Help(I) Then Found = True
Loop

If (Not Found) Then
        MsgBox "word not found in Help document, sorry!", "Ooops!"
    Else
        MsgBox "word exists in document! Hurray!", "Lucky Day"
    End If
End Sub

