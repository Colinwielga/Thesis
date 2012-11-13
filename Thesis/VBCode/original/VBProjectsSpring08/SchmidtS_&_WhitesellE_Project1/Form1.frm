VERSION 5.00
Begin VB.Form frmFrance 
   BackColor       =   &H00FFFFFF&
   Caption         =   "France"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Click Here to See Projected Budget"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Click Here to Display Program Details"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2175
      Left            =   4320
      TabIndex        =   3
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   960
      ScaleHeight     =   2115
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.PictureBox picCannes 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   5400
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lblFrance 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "France"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "frmFrance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGoBack_Click()
frmFrance.Hide
frmEurope.Show
End Sub


