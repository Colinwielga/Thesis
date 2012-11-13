VERSION 5.00
Begin VB.Form GoodNight 
   BackColor       =   &H00400000&
   Caption         =   "Good Night"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit3 
      BackColor       =   &H8000000D&
      Caption         =   "Quit"
      Height          =   495
      Left            =   3600
      MaskColor       =   &H00400000&
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   1440
      Picture         =   "frmGoodNight.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Have a Safe Trip Home"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Please Remember to Turn Off all Lights and Lock all Doors Before you Leave for the Night!"
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "GoodNight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SalesReport (VBProject.vbp)
'Good Night (frmGoodNight.frm)
'Written by Lisa Harren
'10-23-03 (Finished)
'Form Purpose:  Closing screen which allows
      'the user to exit the program.
      

Private Sub cmdQuit3_Click()
'Close the Program

End

End Sub
