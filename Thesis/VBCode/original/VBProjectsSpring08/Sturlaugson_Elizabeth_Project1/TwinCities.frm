VERSION 5.00
Begin VB.Form frmTwinCities 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   12000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12000
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClick 
      BackColor       =   &H80000009&
      Caption         =   "Click to Enter"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   4200
      MaskColor       =   &H80000006&
      TabIndex        =   0
      Top             =   8760
      Width           =   5175
   End
   Begin VB.Image imgIntro 
      Height          =   8805
      Left            =   1680
      Picture         =   "Twin Cities.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   10560
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   3960
      Picture         =   "Twin Cities.frx":456E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5760
   End
End
Attribute VB_Name = "frmTwinCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSCI VB Project: Big Bowl
'frmTwinCities
'Elizabeth K. Sturlaugson
'Due Date: Friday, March 28th, 2008

'This form is the introduction to my project.  Big Bowl is one of my favorite restaurants and I wanted to create a program that allowed the user to perform
'different related tasks that might be of interest to someone who enjoys this restaurant.  The purpose of the project is to educate and to explore some options
'that are related to the restaurant industry, like ordering food and making reservations


Option Explicit

Private Sub cmdClick_Click()
'moves to another form
frmChinese.Show
frmTwinCities.Hide


End Sub

