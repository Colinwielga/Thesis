VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form6"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   7095
      Left            =   7200
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "See Report"
      Height          =   975
      Left            =   7440
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox picCorRep 
      Height          =   7095
      Left            =   480
      ScaleHeight     =   7035
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
