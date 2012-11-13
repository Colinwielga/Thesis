VERSION 5.00
Begin VB.Form Africa 
   BackColor       =   &H00000000&
   Caption         =   "Africa"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Text            =   "Find out more about the Black Continent!"
      Top             =   0
      Width           =   12735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   3735
   End
   Begin VB.CommandButton cmdPopulation 
      Caption         =   "Population"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   3855
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   4200
      Picture         =   "Africa.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   8535
   End
End
Attribute VB_Name = "Africa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: Africa.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  to give the user the ability to switch forms
Option Explicit
'hide the Africa form and then show the main form
Private Sub cmdBack_Click()
Africa.Hide
Main.Show
End Sub
'Hides the Africa Form and Shows the Africa2 Form
Private Sub cmdPopulation_Click()
Africa.Hide
Africa2.Show
End Sub
'Hides the Africa Form and Shows the Africa1 Form
Private Sub cmdSort_Click()
Africa.Hide
Africa1.Show
End Sub



