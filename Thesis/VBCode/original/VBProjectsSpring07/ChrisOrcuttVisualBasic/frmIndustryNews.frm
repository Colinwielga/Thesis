VERSION 5.00
Begin VB.Form frmIndustryNews 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Industry News"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "OCR A Extended"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuitarHero 
      Caption         =   "Guitar Hero:Next?"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdHalo3 
      Caption         =   "Halo 3 News"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   6420
      Left            =   0
      Picture         =   "frmIndustryNews.frx":0000
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "frmIndustryNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmIndustryNews
'26 March 2007

Option Explicit
Private Sub cmdGuitarHero_Click()
    frmIndustryNews.Hide    'Hides IndustryNews form
    frmGuitarHero.Show      'Shows GuitarHero form
End Sub

Private Sub cmdHalo3_Click()
    frmIndustryNews.Hide    'Hides IndustryNews form
    frmHalo3.Show           'Shows Halo3 form
End Sub
Private Sub cmdReturn_Click()
    frmIndustryNews.Hide    'Hides IndustryNews form
    frmSelectWant.Show      'Shows SelectWant form
End Sub
