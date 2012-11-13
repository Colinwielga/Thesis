VERSION 5.00
Begin VB.Form frmXbox360Premium 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Xbox 360 Premium Unit"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   5880
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      Height          =   4815
      Left            =   3720
      ScaleHeight     =   4755
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   840
      Width           =   4215
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lbl360Prem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Xbox 360"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   240
      Picture         =   "frmXbox360Premium.frx":0000
      Top             =   840
      Width           =   2820
   End
End
Attribute VB_Name = "frmXbox360Premium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdReturn_Click()
    frmXbox360Premium.Hide
    frmConsoleInfo.Show
End Sub
