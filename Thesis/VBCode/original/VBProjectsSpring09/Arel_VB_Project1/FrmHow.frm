VERSION 5.00
Begin VB.Form FrmHow 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8295
      Left            =   0
      Picture         =   "FrmHow.frx":0000
      ScaleHeight     =   8235
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "Return to Trivia Menu"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5760
         Width           =   3735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9360
         Picture         =   "FrmHow.frx":13D4D2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6720
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Return to Main"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7920
         Picture         =   "FrmHow.frx":13EBA0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "3. Have FUN! and GOOD LUCK!"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         TabIndex        =   9
         Top             =   3960
         Width           =   7455
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   $"FrmHow.frx":14026E
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1560
         TabIndex        =   8
         Top             =   3000
         Width           =   7455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1. Click the Start Trivia Button"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   7455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "How to Play"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   7455
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Height          =   3975
         Left            =   1200
         TabIndex        =   4
         Top             =   1200
         Width           =   8175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00400000&
         Height          =   4695
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   8895
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1. Click the Start Trivia Button"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "FrmHow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmHow
'Project By: Stephanie Arel
'Date Written: 3/16/2009
'This Form gives a brief "how to" of the trivia game.
Option Explicit

Private Sub Command1_Click()
'Takes the user back to the main Trivia menu
FrmHow.Hide
FrmTrivia.Show
End Sub

Private Sub Command3_Click()
'Ends the program
End
End Sub

Private Sub Command4_Click()
'Takes the user back to the main menu
FrmHow.Hide
FrmMain.Show
End Sub

Private Sub Picture1_Click()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
