VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "The Campground!"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   Picture         =   "frmWelcome.frx":0000
   ScaleHeight     =   5325
   ScaleWidth      =   6360
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000040C0&
      Caption         =   "Goodbye! Have a nice day!"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnterStore 
      BackColor       =   &H000040C0&
      Caption         =   "Come on in!"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      MaskColor       =   &H000080FF&
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Welcome to the Campground, your number one source for camping gear and equipment!"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1935
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEnterStore_Click()
'brings you to the next page/into the store
frmWelcome.Hide
frmHelper.Show
End Sub

Private Sub cmdQuit_Click()
'ends the program
End
End Sub

