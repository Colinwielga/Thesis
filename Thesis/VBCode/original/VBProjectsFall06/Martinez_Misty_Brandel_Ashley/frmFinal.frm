VERSION 5.00
Begin VB.Form frmFinal 
   BackColor       =   &H000000C0&
   Caption         =   "Congratulations!!!!"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label lblDirections 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Click on the picture box below to see how you did!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   6975
   End
   Begin VB.Image imgQuit 
      Height          =   705
      Left            =   3600
      Picture         =   "frmFinal.frx":0000
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   840
      Picture         =   "frmFinal.frx":0505
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   5835
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgQuit_Click()
    End         'ends program
End Sub

Private Sub picResults_Click()
    If Yes = True Then          'Shows results from optYes from Welcome Level
        picResults.Print YourName & " You read a lot of Dr. Seuss books and it shows!!!"
    Else                        'Shows results from optNo from Welcome Level
        picResults.Print YourName & " You sure know a lot even though you haven't read any Dr. Seuss books!!!"
    End If
End Sub
