VERSION 5.00
Begin VB.Form frmmadness 
   BackColor       =   &H00000040&
   Caption         =   "Garrett Sohn"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   29.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdtest 
      Caption         =   "Test yourself"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdrecord 
      Caption         =   "Record Holders"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdchampions 
      Caption         =   "Past champions"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdlearn 
      Caption         =   "Learn about the tourney"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtmarchmadness 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1000
      Left            =   1200
      TabIndex        =   1
      Text            =   "March madness"
      Top             =   480
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3120
      Picture         =   "frmmadness.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   1560
      Width           =   2295
   End
End
Attribute VB_Name = "frmmadness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'March Madness (madness.vbp)
'main form (madness.frm)
'Garrett Sohn
'March 24, 2006
'This form is the starting form which connects to all of the other forms that I have. It also quits the project.

Option Explicit
Private Sub cmdchampions_Click()
    frmmadness.Hide
    frmpast.Show
End Sub

Private Sub cmdlearn_Click()
    frmmadness.Hide
    frminfo.Show
End Sub
Private Sub cmdquit_Click()
    End
End Sub

Private Sub cmdrecord_Click()
    frmmadness.Hide
    frmrecords.Show
End Sub

Private Sub cmdtest_Click()
    frmmadness.Hide
    frmtest.Show
End Sub

