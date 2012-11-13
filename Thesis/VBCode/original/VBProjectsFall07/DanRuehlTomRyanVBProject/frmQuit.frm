VERSION 5.00
Begin VB.Form frmQuit 
   BackColor       =   &H00800000&
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   1080
      Picture         =   "frmQuit.frx":0000
      Top             =   4200
      Width           =   12885
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Thank's for playing Jeopardy! Unfortunately you did not win any money.  Better luck next time!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   3480
      TabIndex        =   0
      Top             =   1560
      Width           =   8295
   End
End
Attribute VB_Name = "frmQuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
