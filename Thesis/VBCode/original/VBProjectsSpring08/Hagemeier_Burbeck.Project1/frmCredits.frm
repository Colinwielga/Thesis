VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H8000000D&
   Caption         =   "Credits"
   ClientHeight    =   12780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18450
   LinkTopic       =   "Form1"
   ScaleHeight     =   12780
   ScaleWidth      =   18450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCredits 
      Caption         =   "See Credits"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   7200
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6915
      ScaleWidth      =   11115
      TabIndex        =   1
      Top             =   120
      Width           =   11175
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Image imgAlps 
      Height          =   18000
      Left            =   -5400
      Picture         =   "frmCredits.frx":0000
      Top             =   -360
      Width           =   24000
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmCredits (Credits)
'Author: Nate Burbeck
'Date Written: 27 March 2008
'Objective: show the user documentation of sources used for this project

Option Explicit
Dim Credits(1 To 100) As String                     'sets Credits as a string

Private Sub cmdCredits_Click()
    Dim CTR As Integer                              'CTR set to 1
    CTR = 1
    Open App.Path & "\credits.txt" For Input As #1  'loads credits.txt
    
    Do While Not EOF(1)
        Input #1, Credits(CTR)                      'names array 'Credits' which was set to string
        CTR = CTR + 1
    Loop
    
    Close #1
    
    Dim i As Integer
    For i = 1 To CTR
        picResults.Print Credits(i)                 'prints array
    Next i
End Sub

Private Sub cmdMainMenu_Click()
frmCredits.Hide                                     'hides this form
frmMainMenu.Show                                    'shows main menu
End Sub

