VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   480
      TabIndex        =   7
      Text            =   "Crime Awareness Project"
      Top             =   0
      Width           =   10695
   End
   Begin VB.CommandButton cmdPark 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check/Pay Parking Tickets"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdNews 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criminals in the News"
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdCops 
      BackColor       =   &H00FFFFFF&
      Caption         =   """Cops"" TV Schedule"
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdBAC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BAC Calculator"
      Height          =   735
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdSexoff 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Sex Offenders In Your Neighborhood?"
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdCollege 
      BackColor       =   &H000000FF&
      Caption         =   "Are You Fit for Law Enforcement?"
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdLicense 
      BackColor       =   &H000000FF&
      Caption         =   "Criminal License Look-Up"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5850
      Left            =   2640
      Picture         =   "HOMEPAGE.frx":0000
      Top             =   1440
      Width           =   5940
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Crime Awareness Project
'frm License
'By: Alex Warzecha & Andrew Enzler
'Written: 3/18/2009
'The objective of this project is to provide the user with a further understanding of the
'world of crimes.
'This form is the main menu for the form. It Contains command buttons that link the
'rest of the program pages.


Private Sub cmdBAC_Click()
'these codes hide a form of choice and brings up another in its place
frmHome.Hide
frmBACalc.Show
End Sub

Private Sub cmdCollege_Click()
frmHome.Hide
frmQuiz.Show
End Sub

Private Sub cmdCops_Click()
frmHome.Hide
frmCops.Show
End Sub

Private Sub cmdLicense_Click()
frmHome.Hide
frmLicense.Show
End Sub

Private Sub cmdNews_Click()
frmHome.Hide
frmNews.Show
End Sub

Private Sub cmdPark_Click()
frmHome.Hide
frmParking.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdPlay_Click()
frmHome.Hide
frmPlayGame.Show
End Sub

Private Sub cmdSexoff_Click()
frmHome.Hide
frmSex.Show
End Sub

Private Sub quit_Click()
End
End Sub
