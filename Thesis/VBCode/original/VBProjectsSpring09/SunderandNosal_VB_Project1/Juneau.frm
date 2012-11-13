VERSION 5.00
Begin VB.Form frmJuneau 
   BackColor       =   &H00000000&
   Caption         =   "Juneau"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   5640
      Picture         =   "Juneau.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton cmdGoBacktoCruisePorts 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Cruise Ports Page"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label ldhjkdf 
      BackColor       =   &H80000009&
      Caption         =   $"Juneau.frx":2EAD
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label lbl34 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Juneau"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmJuneau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmJuneau
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This form is like the last (Ketchikan), but it is basically just giving information instead of having
'the user enter their information into any text boxes.

Private Sub cmdGoBacktoCruisePorts_Click()
frmJuneau.Hide
frmCruisePorts2.Show
End Sub
