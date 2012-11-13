VERSION 5.00
Begin VB.Form frmBooks 
   BackColor       =   &H00000000&
   Caption         =   "Books"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   Picture         =   "frmBooks.frx":0000
   ScaleHeight     =   5415
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBreakingDawn 
      BackColor       =   &H000000C0&
      Caption         =   "Click to learn more about Breaking Dawn"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdEclipse 
      BackColor       =   &H000000C0&
      Caption         =   "Click to learn more about Eclipse"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewMoon 
      BackColor       =   &H000000C0&
      Caption         =   "Click to learn more about New Moon"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdTwilight 
      BackColor       =   &H000000C0&
      Caption         =   "Click to learn more about Twilight"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Twilight
'Form Name: frmBooks
'Author: Mollie Land
'Date Written: 3/15/2009
'Objective: to give the user the option of which book they would like to learn more about

Private Sub cmdBreakingDawn_Click()
    'show Breaking Dawn form, hiding the books form
    frmBreakingDawn.Show
    frmBooks.Hide
End Sub

Private Sub cmdEclipse_Click()
    'show Eclipse form, hiding the books form
    frmEclipse.Show
    frmBooks.Hide
End Sub

Private Sub cmdNewMoon_Click()
    'show New Moon form, hiding the books form
    frmNewMoon.Show
    frmBooks.Hide
End Sub

Private Sub cmdReturn_Click()
    'show main menu, hiding the books form
    frmStart.Show
    frmBooks.Hide
End Sub

Private Sub cmdTwilight_Click()
    'show Twilight form, hiding the books form
    frmTwilight.Show
    frmBooks.Hide
End Sub
