VERSION 5.00
Begin VB.Form frmmainmenu 
   BackColor       =   &H00800000&
   Caption         =   " Main Menu; Project by Kayla Nelson"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmeetauthors 
      BackColor       =   &H008080FF&
      Caption         =   "Meet the Authors"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   2175
   End
   Begin VB.PictureBox piccity 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   6960
      Picture         =   "MainMenu.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   1500
      TabIndex        =   8
      Top             =   1680
      Width           =   1530
   End
   Begin VB.PictureBox pickite 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   4920
      Picture         =   "MainMenu.frx":1CA1
      ScaleHeight     =   2385
      ScaleWidth      =   1545
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.PictureBox picfirst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   2640
      Picture         =   "MainMenu.frx":A51A
      ScaleHeight     =   2625
      ScaleWidth      =   1785
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox picmillion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   240
      Picture         =   "MainMenu.frx":C188
      ScaleHeight     =   2865
      ScaleWidth      =   1905
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H008080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdbuybooks 
      BackColor       =   &H008080FF&
      Caption         =   "Buy the Books"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdreadreviews 
      BackColor       =   &H008080FF&
      Caption         =   "Read Reviews"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdexcerpts 
      BackColor       =   &H008080FF&
      Caption         =   "Read Excerpts"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00800000&
      Caption         =   "Kayla's Book Club"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   975
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmmainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Kayla's Book Club (MainMenu.vbp)
'Form Name: Main Menu (frmmainmenu.frm)
'Author: Kayla Nelson
'Date: 10-27-05
'purpose of the form: To allow the user to select a different form they would like to view.  It also gives them the first look at the books used throughout the program, and also allows them to end the program.

Option Explicit

Private Sub cmdbuybooks_Click() 'This closes the mainmenu and brings up the form to buy the books.
    frmmainmenu.Hide
    frmbuybooks.Show
End Sub

Private Sub cmdexcerpts_Click() 'This closes the main menu and brings up the form to read the excerpts.
    frmmainmenu.Hide
    frmreadexcerpts.Show
End Sub

Private Sub cmdexit_Click() ' This closes the program.
    End
End Sub

Private Sub cmdmeetauthors_Click() ' This closes the main menu form and brings up the form where you can learn about the authors.
    frmmainmenu.Hide
    frmmeetauthors.Show
End Sub


Private Sub cmdreadreviews_Click() 'This closes the main menu and brings up the form to read reviews on the books.
    frmmainmenu.Hide
    frmreadreviews.Show
End Sub
