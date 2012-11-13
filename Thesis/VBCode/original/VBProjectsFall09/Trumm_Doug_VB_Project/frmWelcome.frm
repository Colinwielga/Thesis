VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00004080&
   Caption         =   "The Corporation"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13800
   FillColor       =   &H00C0E0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   13800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H00800000&
      Caption         =   "enter the Work Station"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H00800000&
      TabIndex        =   2
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   3975
   End
   Begin VB.PictureBox picResults 
      Height          =   7215
      Left            =   4440
      Picture         =   "frmWelcome.frx":0000
      ScaleHeight     =   7155
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   1800
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004080&
      Caption         =   "Well that's enough chit-chat.  Let's make some money.  Please proceed to your virtual work station."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004080&
      Caption         =   $"frmWelcome.frx":AFFB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label lblGame 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Welcome to THE CORPORATION"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   13575
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Provide an opening interface with a picture and some explanation


Private Sub cmdSwitch_Click()
    'Moves on to next form
    frmWorkStation.Show
    frmWelcome.Hide
End Sub

