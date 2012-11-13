VERSION 5.00
Begin VB.Form frmIreland 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ireland"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H0080FF80&
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
      TabIndex        =   5
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H0080FF80&
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2655
      Left            =   8760
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   5160
      ScaleHeight     =   2595
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   1320
      Width           =   3735
   End
   Begin VB.PictureBox picIreland 
      Height          =   2655
      Left            =   720
      Picture         =   "frmIreland.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label lblIreland 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ireland"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmIreland"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

End Sub
