VERSION 5.00
Begin VB.Form frmbeginning 
   BackColor       =   &H00400040&
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   FillColor       =   &H00400040&
   ForeColor       =   &H00400040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      TabIndex        =   4
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
   End
   Begin VB.PictureBox picbrettleft 
      Height          =   4575
      Left            =   8400
      Picture         =   "frmbeginning.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   1800
      Width           =   4575
   End
   Begin VB.PictureBox picbrettmid 
      Height          =   4695
      Left            =   240
      Picture         =   "frmbeginning.frx":3FEF
      ScaleHeight     =   4635
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00400040&
      Caption         =   "Welcome to the Brett Favre Super Fan Club:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   10335
   End
End
Attribute VB_Name = "frmbeginning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Brett Favre Fan Club
'Form Name: frmbeginning
'Author: Kory LaCroix
'Date Written: 10/19/08
'Objective: To begin the program
'the purpose of the project is to take the user on an adventure preparing to be a Brett Favre super fan.
'it will do this by enabling the user to buy gear, tickets, plan tickets, and make hotel reservations for away games


Private Sub cmdEnter_Click()
'this moves from this form to the next form
'this begins the program
frmbeginning.Hide
frmgear.Show
End Sub

Private Sub cmdQuit_Click()
'for those who wish, this will end the program
End
End Sub

