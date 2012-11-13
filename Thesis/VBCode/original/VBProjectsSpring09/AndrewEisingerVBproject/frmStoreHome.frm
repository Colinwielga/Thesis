VERSION 5.00
Begin VB.Form frmStoreHome 
   BackColor       =   &H000000FF&
   Caption         =   "Store Home"
   ClientHeight    =   11955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   FillColor       =   &H000000FF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   Picture         =   "frmStoreHome.frx":0000
   ScaleHeight     =   11955
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00004000&
      Caption         =   "Quit"
      Height          =   1575
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9360
      Width           =   2655
   End
   Begin VB.CommandButton cmdReebok 
      BackColor       =   &H00004000&
      Caption         =   "Reebok"
      Height          =   1575
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9360
      Width           =   2655
   End
   Begin VB.CommandButton cmdPuma 
      BackColor       =   &H00004000&
      Caption         =   "Puma"
      Height          =   1575
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdidas 
      BackColor       =   &H00004000&
      Caption         =   "Adidas"
      Height          =   1575
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   2655
   End
   Begin VB.CommandButton cmdNike 
      BackColor       =   &H00004000&
      Caption         =   "Nike"
      Height          =   1575
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   2655
   End
End
Attribute VB_Name = "frmStoreHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program lets the user select a company


Private Sub cmdAdidas_Click()
frmAdidas1.Show
frmStoreHome.Hide
End Sub

Private Sub cmdNike_Click()
frmNike1.Show
frmStoreHome.Hide
End Sub

Private Sub cmdPuma_Click()
frmPuma1.Show
frmStoreHome.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReebok_Click()
frmReebok1.Show
frmStoreHome.Hide
End Sub
