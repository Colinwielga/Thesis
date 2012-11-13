VERSION 5.00
Begin VB.Form frmHelper 
   BackColor       =   &H000040C0&
   Caption         =   "The Campground!"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Leave the store."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   4
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No thanks."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      TabIndex        =   3
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00404080&
      Caption         =   "Yes, please!"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.PictureBox picWendy 
      Height          =   3615
      Left            =   2400
      Picture         =   "frmHelper.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblMeetWendy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hi! I'm Wendy. Can I help you find something?"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNo_Click()
'takes you to the Browse form using the Visible property
frmHelper.Hide
frmBrowse.Show
End Sub

Private Sub cmdQuit_Click()
'ends the program
End
End Sub

Private Sub cmdYes_Click()
'takes you to the ProductSearch form using the Visible property
frmHelper.Hide
frmProductSearch.Show
End Sub
