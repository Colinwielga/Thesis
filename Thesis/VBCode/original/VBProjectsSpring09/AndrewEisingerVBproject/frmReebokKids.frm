VERSION 5.00
Begin VB.Form frmReebokKids 
   Caption         =   "ReebokKids"
   ClientHeight    =   11895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   Picture         =   "frmReebokKids.frx":0000
   ScaleHeight     =   11895
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      Height          =   1335
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdBackGo 
      BackColor       =   &H0080C0FF&
      Caption         =   "Back To Reebok Home"
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FF80&
      Caption         =   "Back To Store Home"
      Height          =   1335
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter A Percent"
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   1215
      Left            =   5880
      OleObjectBlob   =   "frmReebokKids.frx":14D28A
      SourceDoc       =   "M:\CS130\AndrewEisingerVBproject\NBA_on_TNT.mp3"
      TabIndex        =   1
      Top             =   10440
      Width           =   1815
   End
End
Attribute VB_Name = "frmReebokKids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' AthleticStore
' ReebokKids
' Andrew Eisinger
' 3/23/09
'This program gets the input via a inputbox
'This program then message boxes the user

Private Sub cmdBack_Click()
frmStoreHome.Show
frmReebokKids.Hide
End Sub

Private Sub cmdBackGo_Click()
frmReebok1.Show
frmReebokKids.Hide
End Sub

Private Sub cmdEnter_Click()
Dim Percent As Single
Percent = InputBox("Please enter the percent of money you spend on sports items a year.")
    Select Case Percent
        Case Is < 0.1
            MsgBox ("Wow not a big spender!")
        Case Is <= 0.1
            MsgBox ("Getting Better")
        Case Is <= 0.2
            MsgBox ("I remember when I had my first Zima!")
        Case Is <= 0.3
            MsgBox ("My dad spends more and he doesn't play sports!")
        Case Is <= 0.4
            MsgBox ("Really?")
        Case Is <= 0.5
            MsgBox ("You can do better than that!")
        Case Is <= 0.6
            MsgBox ("Buy me a glove cheapo!")
        Case Is <= 0.7
            MsgBox ("Almost there!")
        Case Is <= 0.8
            MsgBox ("Sweet!")
        Case Is >= 0.9
            MsgBox ("Thatta Guy!")
        Case Else
            MsgBox ("Please enter a decimal percentage!")
    
    
End Select
End Sub

Private Sub cmdQuit_Click()
End
End Sub
