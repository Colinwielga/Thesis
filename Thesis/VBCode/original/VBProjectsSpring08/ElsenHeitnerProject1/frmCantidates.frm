VERSION 5.00
Begin VB.Form frmCantidates 
   Caption         =   "Cantidates"
   ClientHeight    =   6165
   ClientLeft      =   3750
   ClientTop       =   2940
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   Picture         =   "frmCantidates.frx":0000
   ScaleHeight     =   6165
   ScaleWidth      =   8535
   Begin VB.TextBox txtDisclaimer 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "All of the information was gathered from each cantidate's official website"
      Top             =   5760
      Width           =   5175
   End
   Begin VB.CommandButton cmdHuckabee 
      Caption         =   "Huckabee's Bio"
      Height          =   615
      Left            =   6000
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdMcCain 
      Caption         =   "McCain's Bio"
      Height          =   615
      Left            =   6000
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdClinton 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clinton's Bio"
      Height          =   615
      Left            =   2040
      MaskColor       =   &H00808080&
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdObama 
      Caption         =   "Obama's Bio"
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox picJohn 
      Height          =   1935
      Left            =   4200
      Picture         =   "frmCantidates.frx":3186
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.PictureBox picBarack 
      Height          =   1935
      Left            =   240
      Picture         =   "frmCantidates.frx":40BC
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox picMike 
      Height          =   1935
      Left            =   4200
      Picture         =   "frmCantidates.frx":50E8
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox PicHillary 
      Height          =   1935
      Left            =   240
      Picture         =   "frmCantidates.frx":5F1A
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Intro "
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "frmCantidates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: Cantidates(frmCantidates.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 12, 2008
'PURPOSE:  This form is our cantidate platform.  Here you can choose which Bio to go to and learn about that cantidate.

Option Explicit

'Takes you back to the intro form
Private Sub cmdBack_Click()
frmCantidates.Hide
frmIntro.Show
End Sub

'Takes you to the Biography of Hillary Clinton
Private Sub cmdClinton_Click()
frmCantidates.Hide
frmClinton.Show
End Sub

'Takes you to the Biography of Mike Huckabee
Private Sub cmdHuckabee_Click()
frmCantidates.Hide
frmHuckabee.Show
End Sub

'Takes you to the Biography of John McCain
Private Sub cmdMcCain_Click()
frmCantidates.Hide
frmMcCain.Show
End Sub

'Takes you to the Biography of Borack Obama
Private Sub cmdObama_Click()
frmCantidates.Hide
frmObama.Show
End Sub

