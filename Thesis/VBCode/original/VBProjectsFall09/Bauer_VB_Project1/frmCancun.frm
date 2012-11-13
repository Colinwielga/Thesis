VERSION 5.00
Begin VB.Form frmCancun 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton cmdCampriRiviera 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Campri Riviera"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   2775
   End
   Begin VB.CommandButton cmdOmni 
      BackColor       =   &H0080C0FF&
      Caption         =   "Omni Hotel"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Label lblCampri 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   $"frmCancun.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   3
      Top             =   6600
      Width           =   4215
   End
   Begin VB.Label lblOmni 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "This Is the Omni Hotel, It has a great party sceen. It is a little less expensive then the Campri Riviera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label lblWhereToStay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Where would you like to stay?"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   9015
   End
   Begin VB.Image Image2 
      Height          =   3360
      Left            =   5160
      Picture         =   "frmCancun.frx":008D
      Top             =   3000
      Width           =   4275
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   240
      Picture         =   "frmCancun.frx":2EDCF
      Top             =   2880
      Width           =   4275
   End
   Begin VB.Label lblCancun 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "CANCUN!!!!!!"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "frmCancun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Cancun page'
'This page is a frm to frm transition page'
'i created it to have options and give people choices
'OCtober 14th 2009'
'Blake Bauer'
'hide and show frms'
Private Sub cmdCampriRiviera_Click()

    frmCancun.Hide
    frmHotel.Show
End Sub
'quit button'
Private Sub cmdEnd_Click()
    End
End Sub

'hide and show frms'
Private Sub cmdOmni_Click()
    frmCancun.Hide
    frmHotel.Show
End Sub

