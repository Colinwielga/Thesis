VERSION 5.00
Begin VB.Form frmAlliance 
   BackColor       =   &H0080FFFF&
   Caption         =   "Alliance"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   Picture         =   "frmAlliance.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      TabIndex        =   3
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdBend 
      Caption         =   "I accept his offer."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdDecline 
      Caption         =   "I shall never bend the knee.  We shall fight and my sons will rule the North for centuries to come!"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   8280
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H0080FFFF&
      Caption         =   $"frmAlliance.frx":9815
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   8055
   End
End
Attribute VB_Name = "frmAlliance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with two options: either taking the alliance offered
'or declining it
'the dicsion will affect the boolean variable defining whether or not the option was taken or not
'and if take, the amount of battlepoints the user has will increase by 3000


Private Sub Command5_Click()
End
End Sub

Private Sub cmdBend_Click()
'army increases and battlepoints increase and individual unit points increase
'this opens up the option for a different outcome, so a boolean must be made

LannisterAllianceN = True
'2000 infantry, 300 archers, 20 knights
Battlepoints = Battlepoints + (2000 * 1) + (200 * 4) + (20 * 10)
frmAlliance.Hide
frmCouncilors2.Show
End Sub

Private Sub cmdDecline_Click()
frmAlliance.Hide
frmCouncilors2.Show
End Sub
