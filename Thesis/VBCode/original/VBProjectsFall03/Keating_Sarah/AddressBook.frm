VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Addresses, Phone Numbers, and Email"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pictureDisplayBox 
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   2535
      Left            =   1920
      ScaleHeight     =   2475
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdAddress 
      Caption         =   "Click to look up an address"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
