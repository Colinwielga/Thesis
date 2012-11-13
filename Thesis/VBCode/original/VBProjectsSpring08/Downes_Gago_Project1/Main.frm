VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00008000&
   Caption         =   "Main"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   13110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAntartica 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ANTARCTICA"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8880
      TabIndex        =   5
      Top             =   7200
      Width           =   2535
   End
   Begin VB.TextBox Wheretxt 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Text            =   "Choose Continent That You Would Like to Explore"
      Top             =   120
      Width           =   12015
   End
   Begin VB.CommandButton cmdAsia 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ASIA"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton cmdEurope 
      BackColor       =   &H00C0FFFF&
      Caption         =   "EUROPE"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdAfrica 
      BackColor       =   &H00C0FFFF&
      Caption         =   "AFRICA"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton cmdSouthAmerica 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SOUTH AMERICA"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   2520
      Picture         =   "Main.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   9885
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: Main.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  The objective of this form is to be a starting point
'With buttons to bring the user to other forms throughout our program

Option Explicit
'The Main Form is Hidden and the Africa Form is shown
Private Sub cmdAfrica_Click()
Main.Hide
Africa.Show
End Sub
'The Main Form is Hidden and the Antarctica Form is shown
Private Sub cmdAntartica_Click()
Main.Hide
Antarctica.Show
End Sub
'The Main Form is Hidden and the Asia Form is Shown
Private Sub cmdAsia_Click()
Main.Hide
Asia.Show
End Sub
'The Main Form is Hidden and the Europe Form is Shown
Private Sub cmdEurope_Click()
Main.Hide
Europe.Show
End Sub
'Quits the program
Private Sub cmdQuit_Click()
End
End Sub
'The Main Form is Hidden and the South America Form is Shown
Private Sub cmdSouthAmerica_Click()
Main.Hide
SouthAmerica.Show
End Sub
