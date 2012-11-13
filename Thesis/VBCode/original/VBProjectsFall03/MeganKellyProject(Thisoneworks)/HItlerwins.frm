VERSION 5.00
Begin VB.Form HItlerwins 
   BackColor       =   &H00400000&
   Caption         =   "Rollin' with his homies..."
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form2"
   ScaleHeight     =   8475
   ScaleWidth      =   8550
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton QuitHitlerW 
      BackColor       =   &H00808080&
      Caption         =   "Get me out of here!"
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   8160
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   $"HITLER~1.frx":0000
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   5400
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   4395
      Left            =   1080
      Picture         =   "HITLER~1.frx":00C6
      Top             =   840
      Width           =   6090
   End
End
Attribute VB_Name = "HItlerwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub QuitHitlerW_Click()
End
End Sub
