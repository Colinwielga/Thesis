VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Poor Richard"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   1
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H0000FF00&
      Caption         =   "Click Here To Enter!"
      Height          =   855
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   6960
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   8700
      Left            =   -120
      Picture         =   "frmMain.frx":0000
      Top             =   -120
      Width           =   10845
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR As Integer 'declaring variables

Private Sub cmdEnter_Click()
    Names = InputBox("Enter your name to begin:", "Welcome!") 'having the viewer put thier name into the program
    frmMain.Hide 'hiding the main page
    frmHome.Show 'showing the home page
    MsgBox "Welcome to Jonnie Wrestling " & Names & "!" 'creating message
End Sub


Private Sub cmdQuit_Click()
    End 'ending the program
End Sub


