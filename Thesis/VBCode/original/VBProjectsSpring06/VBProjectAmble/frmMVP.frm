VERSION 5.00
Begin VB.Form frmMVP 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdround6 
      Caption         =   "Championship"
      Height          =   855
      Left            =   6480
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdround5 
      Caption         =   "Final 4"
      Height          =   855
      Left            =   3480
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdround4 
      Caption         =   "Elite 8"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdround3 
      Caption         =   "Sweet 16"
      Height          =   855
      Left            =   6480
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdround2 
      Caption         =   "2nd Round"
      Height          =   855
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdround1 
      Caption         =   "1st Round"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lbljeff 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Jeff Amble"
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Image imgmvp 
      Height          =   1815
      Left            =   1440
      Picture         =   "frmMVP.frx":0000
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label lblpickdate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on the round of the game you wish to research"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmMVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form enables the user to search for the MVP of any team'
'in any round of competition'
Option Explicit
'This button enables the user to go back to the main page'
Private Sub cmdback_Click()
    frmMVP.Visible = False
    frmmain.Visible = True
End Sub
'This button enables the user to go to a form for the'
'corresponding round'
Private Sub cmdround1_Click()
    frmMVP.Visible = False
    frmMVPround1.Visible = True
End Sub
'This button enables the user to go to a form for the'
'corresponding round'
Private Sub cmdround2_Click()
    frmMVP.Visible = False
    frmMVPround2.Visible = True
End Sub
'This button enables the user to go to a form for the'
'corresponding round'
Private Sub cmdround3_Click()
    frmMVP.Visible = False
    frmMVPround3.Visible = True
End Sub
'Private'This button enables the user to go to a form for the'
'corresponding round'
Sub cmdround4_Click()
    frmMVP.Visible = False
    frmMVPround4.Visible = True
End Sub
'This button enables the user to go to a form for the'
'corresponding round'
Private Sub cmdround5_Click()
    frmMVP.Visible = False
    frmMVPround5.Visible = True
End Sub
'This button enables the user to go to a form for the'
'corresponding round'
Private Sub cmdround6_Click()
    frmMVP.Visible = False
    frmMVPround6.Visible = True
End Sub
