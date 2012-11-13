VERSION 5.00
Begin VB.Form frmregister 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   Picture         =   "frmregister.frx":0000
   ScaleHeight     =   6165
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfood 
      Caption         =   "Food Service"
      Height          =   855
      Left            =   4920
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdregister 
      Caption         =   "Register Now!"
      Height          =   975
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Go Back to Majorlist"
      Height          =   855
      Left            =   4800
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      Height          =   5775
      Left            =   360
      ScaleHeight     =   5715
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmregister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Written by Yuzu Takei
' Written 10-17-09
Dim Major As String, CN As Integer, Classes As String, Credits As Integer
Private Sub cmdgoback_Click()
'it allows the user to go back to majorlist form
frmmajrlist.Show
End Sub
Private Sub cmdregister_Click()
Dim Totalcredits As Integer
'it asks the user to put major, class name, class number, and credits in the input box
picresults.Print "Major", "Class No.", "Claa Name", "Credits"
picresults.Print "***************************************************************"
Totalcredits = 0
Do Until Totalcredits >= 12 And Credits <= 18
    Totalcredits = Totalcredits + Credits
    Major = InputBox("Enter Major(Capitalized)", "Major Entry")
    CN = InputBox("Enter Class Number", "Number Entry")
    Classes = InputBox("Enter Class Name", "Name Entry")
    Credits = InputBox("Enter Credits", "Credits Entry")
    picresults.Print Major, CN, Classes, Credits
Loop
MsgBox "Congratulations, You have successfully registered for Fall 2009!"
End Sub

Private Sub cmdfood_Click()
'to go to next form
frmfoodservice.Show
frmregister.Hide
End Sub
