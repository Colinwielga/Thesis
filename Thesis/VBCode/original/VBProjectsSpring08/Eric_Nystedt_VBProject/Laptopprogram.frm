VERSION 5.00
Begin VB.Form frmStartup 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Welcome!!!!! Ready to Shop???"
   ClientHeight    =   13425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19815
   LinkTopic       =   "Form1"
   ScaleHeight     =   13425
   ScaleWidth      =   19815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cndquit1 
      BackColor       =   &H000000FF&
      Caption         =   "Click Here to Exit"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   4095
   End
   Begin VB.CommandButton cmdenter 
      BackColor       =   &H0000C000&
      Caption         =   "Click Here View Laptops"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   4770
      Left            =   4440
      Picture         =   "Laptopprogram.frx":0000
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: laptop data analysis
'form: frmstartup
'Written by : Eric Nystedt
'March 22, 2008
'The purpose of this program is to let the user selecet a variety of laptops and the arrange them based on some of their features.
'this specific form is the entrance page to the program


Option Explicit
' This button allows the user to switch to the form that is the base of the program
Private Sub cmdenter_Click()
frmMainPage.Visible = True
frmStartup.Visible = False


End Sub

Private Sub cndquit1_Click()
End
End Sub

