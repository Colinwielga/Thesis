VERSION 5.00
Begin VB.Form frmSecond 
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   Picture         =   "frmSecond.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000013&
      Caption         =   "Go Back"
      Height          =   1095
      Left            =   4560
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.PictureBox picResult 
      Height          =   2175
      Left            =   480
      Picture         =   "frmSecond.frx":FB34
      ScaleHeight     =   2115
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label lblWhat 
      BackColor       =   &H80000013&
      Caption         =   $"frmSecond.frx":120A5
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmSecond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Chen,Guo,Shi,Tian_Project1
'Form Name: frmSecond
'Author: Chen, Zhongjie
        'Guo, Zhishan
        'Shi, Yimei
        'Tian, Yukun
'Date Written: Oct. 22
'Objective: This form shows the rules of NIM

Private Sub cmdBack_Click()
frmSecond.Hide
frmFirst.Show
End Sub

