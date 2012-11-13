VERSION 5.00
Begin VB.Form Metrodome 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Metrodome Vs. New Stadium"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Metrodome.frx":0000
      Top             =   600
      Width           =   8775
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Return To Homepage"
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   480
      Picture         =   "Metrodome.frx":0201
      ScaleHeight     =   2955
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   4800
      Width           =   3855
   End
   Begin VB.PictureBox Picture2 
      Height          =   4095
      Left            =   4440
      Picture         =   "Metrodome.frx":24B0
      ScaleHeight     =   4035
      ScaleWidth      =   5955
      TabIndex        =   3
      Top             =   3960
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Proposed Twins Ballpark"
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Metrodome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project Name: Twins Baseball
' Form Name: Metrodome
' Authors: Jake Krisnik & Mike Foley
' Date Written: October 24, 2006
' Form Objective: To provide the user with information regarding the proposed new Twins
'                 Ballpark which is planned to open in 2010. The form has a picture of
'                 the current Metrodome and a sketch of what the new open air stadium
'                 will look like.
Option Explicit
Private Sub cmdReturn_Click()
' This command button allows the user to navigate away from the Metrodome form and return
' to the Homepage.
    Metrodome.Hide
    HomePage.Show
End Sub

