VERSION 5.00
Begin VB.Form FrmBarackObama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   1215
      Left            =   6960
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   1215
      Left            =   6960
      TabIndex        =   7
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1215
      Left            =   6960
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1215
      Left            =   4920
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   4920
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   240
      ScaleHeight     =   5295
      ScaleWidth      =   4215
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "About Barack"
      Height          =   1215
      Left            =   4920
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.PictureBox PicBarack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   4680
      Picture         =   "FrmBarackObama.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   4095
      TabIndex        =   1
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label LblBarack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "                                   Barack Obama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "FrmBarackObama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAbout_Click()
PicResults.Cls
PicResults.Print
End Sub

