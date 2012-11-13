VERSION 5.00
Begin VB.Form frmAustria 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Austrian Program"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H00FF8080&
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00FF8080&
      Caption         =   "Click Here to See Program Details"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00C0E0FF&
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
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2655
      Left            =   4680
      TabIndex        =   3
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   960
      ScaleHeight     =   2595
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   5880
      Picture         =   "frmAustria.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblAustria 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Austria"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
End
Attribute VB_Name = "frmAustria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGoBack_Click()
frmAustria.Hide
frmEurope.Show
End Sub

