VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Western Europe Travel Log"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   Palette         =   "frmStart.frx":0000
   Picture         =   "frmStart.frx":1DCB8
   ScaleHeight     =   6990
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Click here to Enter."
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the Western European Travel Log: Your interactive gateway to Europe."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmUKMAP
'Author: Nate Burbeck
'Date Written: 26 March 2008
'Objective: To have a startup page to allow the user to enter the program
'Overal objective: to allow the user to discover western Europe through Maps, articles, surveys, pictures, and a blog

Private Sub cmdBegin_Click()
frmStart.Hide
frmMainMenu.Show
End Sub

