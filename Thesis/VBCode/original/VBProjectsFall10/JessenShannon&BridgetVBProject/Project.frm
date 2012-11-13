VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMenuResults 
      Height          =   6615
      Left            =   1440
      ScaleHeight     =   6555
      ScaleWidth      =   8475
      TabIndex        =   1
      Top             =   1320
      Width           =   8535
   End
   Begin VB.Label lblMenu 
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
