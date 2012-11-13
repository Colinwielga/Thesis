VERSION 5.00
Begin VB.Form frmPretty 
   BackColor       =   &H00800080&
   Caption         =   "Thelma Beautiful"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "BYE!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      TabIndex        =   5
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdDonate 
      Caption         =   "Donate More!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   7080
      Picture         =   "frmPretty.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Go Home!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7320
      Picture         =   "frmPretty.frx":0DC1
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdShowme 
      Caption         =   "Show Me!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      TabIndex        =   2
      Top             =   4800
      Width           =   2295
   End
   Begin VB.PictureBox samara 
      Height          =   3975
      Left            =   840
      Picture         =   "frmPretty.frx":1BFA
      ScaleHeight     =   3915
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.PictureBox penelope 
      Height          =   5775
      Left            =   1440
      Picture         =   "frmPretty.frx":4ED0C
      ScaleHeight     =   5715
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
End
Attribute VB_Name = "frmPretty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDonate_Click()
'Opens New form

frmPretty.Hide
frmDonate.Show
End Sub

Private Sub cmdHome_Click()
'Opens New form

frmPretty.Hide
frmDoll.Show
End Sub

Private Sub cmdQuit_Click()
'Exits program
End
End Sub

Private Sub cmdShowme_Click()
'Displays a picture depending on how much money was donated
If total > 1000 Then
    penelope.Visible = True
    samara.Visible = False
    MsgBox ("I feel like a completely new person!")
Else: samara.Visible = True
    penelope.Visible = False
    MsgBox ("I think you may need to donate some more money!")
End If
End Sub
