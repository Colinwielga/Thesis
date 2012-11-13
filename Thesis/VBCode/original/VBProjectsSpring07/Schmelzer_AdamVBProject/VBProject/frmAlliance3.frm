VERSION 5.00
Begin VB.Form frmAlliance3 
   BackColor       =   &H0080FFFF&
   Caption         =   "Alliance"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   Picture         =   "frmAlliance3.frx":0000
   ScaleHeight     =   6945
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDecline 
      Caption         =   "I have no need for the this sycophant of the south.  He may be plotting against me."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8280
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
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
      TabIndex        =   2
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton cmdBend2 
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
      Left            =   8160
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H0080FFFF&
      Caption         =   $"frmAlliance3.frx":9815
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   8055
   End
End
Attribute VB_Name = "frmAlliance3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form presents the user with two options: either taking the alliance offered
'or declining it
'the dicsion will affect the boolean variable defining whether or not the option was taken or not
'and if take, the amount of battlepoints the user has will increase by 3000

Private Sub cmdBend2_Click()
LannisterAllianceP = True
'2000 infantry, 300 archers, 20 knights
Battlepoints = Battlepoints + (2000 * 1) + (200 * 4) + (20 * 10)
frmAlliance3.Hide
frmCouncilors2.Show
End Sub

Private Sub cmdDecline_Click()
frmAlliance3.Hide
frmCouncilors2.Show
End Sub
