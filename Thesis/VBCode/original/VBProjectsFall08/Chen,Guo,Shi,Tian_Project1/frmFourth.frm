VERSION 5.00
Begin VB.Form frmFourth 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   Picture         =   "frmFourth.frx":0000
   ScaleHeight     =   4305
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   1575
      Left            =   9600
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblStrategy4 
      BackColor       =   &H80000013&
      Caption         =   $"frmFourth.frx":FB34
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   8895
   End
   Begin VB.Label lblStrategy3 
      BackColor       =   &H80000013&
      Caption         =   $"frmFourth.frx":FC76
      ForeColor       =   &H80000001&
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   8895
   End
   Begin VB.Label lblStrategy2 
      BackColor       =   &H80000013&
      Caption         =   $"frmFourth.frx":FE4F
      ForeColor       =   &H80000001&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   8895
   End
   Begin VB.Label lblStrategy1 
      BackColor       =   &H80000013&
      Caption         =   "QUESTION: Which positions are losing for the player whose turn it is to move? "
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "frmFourth"
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
'Date Written: Oct. 28
'Objective: This form simply show the winning strategy of the game NIM

Private Sub cmdBack_Click()
frmFourth.Visible = False
frmThird.Visible = True
End Sub
