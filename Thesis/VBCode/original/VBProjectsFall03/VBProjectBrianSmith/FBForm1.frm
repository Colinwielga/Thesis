VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Click Here to Get Started"
      Height          =   855
      Left            =   1080
      TabIndex        =   1
      Top             =   3360
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   840
      Picture         =   "FB Form1.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Brian Smith
    'Project1(Football Project)
    'Form2(FB Form2)
    'Oct 26th, 2003
    'Purpose: To calculate the number of points a receiver has acumulated for fantasy football
    


Private Sub cmdStart_Click()
Form1.Visible = False
Form2.Visible = True 'Switches to the second form (form2)

End Sub
