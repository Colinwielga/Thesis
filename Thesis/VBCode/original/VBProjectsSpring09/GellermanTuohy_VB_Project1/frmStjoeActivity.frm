VERSION 5.00
Begin VB.Form frmStjoeActivity 
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   855
      Left            =   9720
      TabIndex        =   0
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   5130
      Left            =   4680
      Picture         =   "frmStjoeActivity.frx":0000
      Top             =   480
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   840
      Picture         =   "frmStjoeActivity.frx":7091A
      Top             =   1560
      Width           =   3510
   End
End
Attribute VB_Name = "frmStjoeActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
    
    frmStjoeActivity.Hide
    frmHotel.Show
    
End Sub
