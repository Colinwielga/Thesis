VERSION 5.00
Begin VB.Form frmChinese 
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   11370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form2"
   ScaleHeight     =   11370
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChinese 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here for more information about Big Bowl!"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      TabIndex        =   0
      Top             =   10920
      Width           =   6735
   End
   Begin VB.Image img3 
      Height          =   2415
      Left            =   2880
      Picture         =   "Choose Restaurant.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3015
   End
   Begin VB.Image img2 
      Height          =   2445
      Left            =   8640
      Picture         =   "Choose Restaurant.frx":0DC4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2985
   End
   Begin VB.Image img1 
      Height          =   2340
      Left            =   14520
      Picture         =   "Choose Restaurant.frx":1972
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Choose Restaurant.frx":2A3A
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro H"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   6960
      TabIndex        =   1
      Top             =   3960
      Width           =   7095
   End
   Begin VB.Image imgUpdate 
      Height          =   11520
      Left            =   2640
      Picture         =   "Choose Restaurant.frx":2B1B
      Top             =   240
      Width           =   15360
   End
End
Attribute VB_Name = "frmChinese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSCI VB Project: Big Bowl
'frmChinese
'Elizabeth K. Sturlaugson
'Due Date: Friday, March 28th, 2008

'The purpose of this form to give the user some basic information about Big Bowl

Option Explicit

Private Sub cmdChinese_Click()
'moves to another form

frmBigBowl.Show
frmChinese.Hide

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Text1_Change()

End Sub

