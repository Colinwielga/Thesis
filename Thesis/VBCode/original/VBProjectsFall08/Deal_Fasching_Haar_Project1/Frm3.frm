VERSION 5.00
Begin VB.Form frmbuttons 
   Caption         =   "Form2"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11970
   LinkTopic       =   "Form2"
   Picture         =   "Frm3.frx":0000
   ScaleHeight     =   9285
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to East High"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10440
      TabIndex        =   5
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdwhich 
      Caption         =   "Which High School Musical Character are You?"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdauthors 
      Caption         =   "Get To Know the Authors!"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdtrivia 
      Caption         =   "Test Your High School Musical Knowledge!"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdtune 
      Caption         =   "Name That Tune!"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdgettoknow 
      Caption         =   "Get to Know The Characters of East High!"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmbuttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdauthors_Click()
frmauthors.Show
frmbuttons.Hide
Frmcharacter.Hide
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Hide
frmtitle.Hide
End Sub

Private Sub cmdback_Click()
frmauthors.Hide
frmbuttons.Hide
Frmcharacter.Hide
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Hide
frmtitle.Show
End Sub

Private Sub cmdgettoknow_Click()
frmauthors.Hide
frmbuttons.Hide
Frmcharacter.Show
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Hide
End Sub

Private Sub cmdtrivia_Click()
frmauthors.Hide
frmbuttons.Hide
Frmcharacter.Hide
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Show
frmquiz.Hide
End Sub

Private Sub cmdtune_Click()
frmauthors.Hide
frmbuttons.Hide
Frmcharacter.Hide
frmtitle.Hide
frmnamethattune.Show
FrmTrivia.Hide
frmquiz.Hide
End Sub

Private Sub cmdwhich_Click()
frmauthors.Hide
frmbuttons.Hide
Frmcharacter.Hide
frmtitle.Hide
frmnamethattune.Hide
FrmTrivia.Hide
frmquiz.Show
End Sub