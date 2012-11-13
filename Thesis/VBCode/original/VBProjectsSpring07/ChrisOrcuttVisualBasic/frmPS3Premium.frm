VERSION 5.00
Begin VB.Form frmPS3 
   BackColor       =   &H00000000&
   Caption         =   "Sony Playstation 3 (60 Gig) Unit"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRetun 
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
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3315
      Left            =   0
      Picture         =   "frmPS3Premium.frx":0000
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmPS3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdRetun_Click()
    frmPS3.Hide
    frmConsoleInfo.Show
End Sub
