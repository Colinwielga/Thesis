VERSION 5.00
Begin VB.Form frmfirstform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmfirstform.frx":0000
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGame 
      BackColor       =   &H0000FFFF&
      Caption         =   "Game"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5273
      Width           =   1935
   End
   Begin VB.PictureBox Picture4 
      Height          =   1695
      Left            =   240
      Picture         =   "frmfirstform.frx":223F4
      ScaleHeight     =   1635
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   9120
      Width           =   2295
   End
   Begin VB.PictureBox Picture3 
      Height          =   1815
      Left            =   12720
      Picture         =   "frmfirstform.frx":22D00
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   12840
      Picture         =   "frmfirstform.frx":23C09
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   8640
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   120
      Picture         =   "frmfirstform.frx":2524D
      ScaleHeight     =   915
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6488
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6353
      Width           =   2295
   End
   Begin VB.CommandButton cmdFinder 
      BackColor       =   &H00008000&
      Caption         =   "Gun Finder"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6668
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4073
      Width           =   1935
   End
   Begin VB.Label lblsponsor 
      BackColor       =   &H000000FF&
      Caption         =   "These are the sponsors of the program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4875
      TabIndex        =   6
      Top             =   3480
      Width           =   5520
   End
End
Attribute VB_Name = "frmfirstform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Gun Selector (Zach Meyer's VB Project.vbp)
'Form Name : frmfirstform (frmfirstform.frm)
'Author: Zach Meyer
'Date Written: October 26, 2005
'Objective: It serves as a main menu for the program
                 'and when the buttons are clicked makes the user
                 'go to another form.
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.

Option Explicit

'This button will advance the user to the second form.

Private Sub cmdFinder_Click()
    frmfirstform.Hide
    frmsecondform.Show
End Sub

'This will exit the program.

Private Sub cmdExit_Click()
    End
End Sub

'This button will send the user to the third form, which is the game.
'It also makes all of the birds on the form appear.

Private Sub cmdGame_Click()
    frmfirstform.Hide
    frmsecondform.Hide
    frmthirdform.Show
End Sub

