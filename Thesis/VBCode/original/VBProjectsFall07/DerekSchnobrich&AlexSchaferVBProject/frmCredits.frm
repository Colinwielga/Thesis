VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H0000C000&
   Caption         =   "Credits+Works Cited"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   FillColor       =   &H0000C000&
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDerek 
      Height          =   3135
      Left            =   4440
      Picture         =   "frmCredits.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picSchaf 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   1920
      Picture         =   "frmCredits.frx":BFB2
      ScaleHeight     =   2835
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdWorks 
      BackColor       =   &H0000FFFF&
      Caption         =   "Works Cited"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdCredits 
      BackColor       =   &H0000FFFF&
      Caption         =   "Credits"
      Height          =   735
      Left            =   240
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   600
      Width           =   9135
   End
   Begin VB.Label lbl2 
      BackColor       =   &H000080FF&
      Caption         =   "Derek Schnobirch"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FF00FF&
      Caption         =   "Alex Schafer"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Prints the Credits
Private Sub cmdCredits_Click()
picResults.Print "The code included in this program was produced by Alex Schafer and Derek Schnobrich."
picResults.Print "Some code features were learned through http://www.vb6.us/"
cmdCredits.Enabled = False
lbl1.Visible = True
lbl2.Visible = True
picSchaf.Visible = True
picDerek.Visible = True
End Sub
'Quits the program
Private Sub cmdQuit_Click()
 End
End Sub

'Returns the user to the main Menu
Private Sub cmdReturn_Click()
frmCredits.Hide
frmHome.Show
End Sub

'Prints the Works Cited
Private Sub cmdWorks_Click()
picResults.Print Chr(10); "The pictures, names, and data included in this program are taken from www.gojohnnies.com and www.d3football.com."; Chr(10); "The writers of this program are not responsible for the dominating tradition of Johnnie Football."; Chr(10); Chr(10); "'We're Just ordianry people, doing ordinary things, extraordinarily well.'"; Chr(10); "-John Gagliardi"
cmdWorks.Enabled = False
End Sub
