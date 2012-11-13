VERSION 5.00
Begin VB.Form frmFind 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Find Right Program"
   ClientHeight    =   6810
   ClientLeft      =   2835
   ClientTop       =   2220
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   10350
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H0080FFFF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdInterests 
      BackColor       =   &H008080FF&
      Caption         =   "Click Here to Find a Program Based on Your Interests"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click Here to Find a Program Based on Your Budget"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   2880
      Picture         =   "frmFind.frx":0000
      ScaleHeight     =   4575
      ScaleWidth      =   4575
      TabIndex        =   2
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   Find the Program that is the                                                               Best Fit For YOU!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   10335
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written 3/11/08 by Sammi and Erika


Private Sub cmdGoBack_Click()
frmFind.Hide
frmPrograms.Show
End Sub

Private Sub cmdBudget_Click()
frmFind.Hide
frmFindBudget.Show
End Sub

Private Sub cmdInterests_Click()
frmFind.Hide
frmFindInterests.Show
End Sub
