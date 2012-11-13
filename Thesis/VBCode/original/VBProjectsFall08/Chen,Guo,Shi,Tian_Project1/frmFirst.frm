VERSION 5.00
Begin VB.Form frmFirst 
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   Picture         =   "frmFirst.frx":0000
   ScaleHeight     =   7050
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicResult1 
      Height          =   5295
      Left            =   360
      Picture         =   "frmFirst.frx":FB34
      ScaleHeight     =   5235
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   1440
      Width           =   5895
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000013&
      Caption         =   "Quit"
      Height          =   735
      Left            =   6720
      Picture         =   "frmFirst.frx":1B1DD
      TabIndex        =   2
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdWhat 
      BackColor       =   &H80000013&
      Caption         =   "Rules of NIM"
      Height          =   735
      Left            =   6720
      Picture         =   "frmFirst.frx":21329
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H80000013&
      Caption         =   "Start the Game"
      Height          =   735
      Left            =   6720
      Picture         =   "frmFirst.frx":27475
      TabIndex        =   0
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label lblNIM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NIM!!!"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Chen,Guo,Shi,Tian_Project1
'Form Name: frmFirst
'Author: Chen, Zhongjie
        'Guo, Zhishan
        'Shi, Yimei
        'Tian, Yukun
'Date Written: Oct. 20
'Objective: This is the startup form of the project. The project
            'is to demonstrate the rules and winning strategies'
            'of the game NIM. Most importantly, it is FUN!

Option Explicit
Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStart_Click()
frmFirst.Hide
frmThird.Show
MsgBox "Welcome to the World of NIM!", , "Welcome!"
End Sub

Private Sub cmdWhat_Click()
frmFirst.Hide
frmSecond.Show
End Sub

Private Sub Command1_Click()
frmFirst.Hide
frmFourth.Show
End Sub
