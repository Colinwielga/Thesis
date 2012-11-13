VERSION 5.00
Begin VB.Form frmTortolla 
   BackColor       =   &H008080FF&
   Caption         =   "Tortolla"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackToCaribbean 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Caribbean Home Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblTortollaInfo 
      BackColor       =   &H00C0C0FF&
      Caption         =   $"Tortolla.frx":0000
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Label lblTortolla 
      BackColor       =   &H00C0C0FF&
      Caption         =   "   Tortola"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmTortolla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmTortolla
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form is like the last (St. Thomas), but it is basically just giving information instead of having
'the user enter their information into any text boxes.

Private Sub cmdBackToCaribbean_Click()
frmTortolla.Hide
frmCaribbeanHome.Show
End Sub
