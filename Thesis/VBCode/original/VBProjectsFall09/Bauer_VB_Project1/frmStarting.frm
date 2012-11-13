VERSION 5.00
Begin VB.Form frmStarting 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit! (Why?)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Click Here To Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   $"frmStarting.frx":0000
      BeginProperty Font 
         Name            =   "Mathematica7"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   9495
   End
   Begin VB.Image Image2 
      Height          =   2910
      Left            =   5280
      Picture         =   "frmStarting.frx":00B2
      Top             =   1440
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   2910
      Left            =   600
      Picture         =   "frmStarting.frx":2AAFC
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Spring Break Adventures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmStarting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'welcom to my project
'I want to start a travel agency'
'this project will help anyone find a spring break destination prefect for them'
'ocotber 14th 2009
'Blake Bauer'

Option Explicit
'Quit Button'
Private Sub cmdEnd_Click()
    End
End Sub

'Show and hide my first frm'
Private Sub cmdStart_Click()
    frmPlace.Show
    frmStarting.Hide
End Sub

