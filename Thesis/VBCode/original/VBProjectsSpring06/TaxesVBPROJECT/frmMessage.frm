VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H80000009&
   Caption         =   "A Message From the Commissioner"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   Picture         =   "frmMessage.frx":0000
   ScaleHeight     =   9315
   ScaleWidth      =   13590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturnFrontpage 
      Caption         =   "Go to Homepage"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblMessage 
      Caption         =   " IRS Commitioner's          Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   3
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H8000000C&
      Caption         =   "By Brent Mergen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   2
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Image ImgLetter 
      Height          =   6990
      Left            =   3720
      Picture         =   "frmMessage.frx":326E3
      Top             =   1560
      Width           =   8805
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Mark W. Everson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3870
      Left            =   360
      Picture         =   "frmMessage.frx":FB22D
      Top             =   240
      Width           =   3090
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E-Z Tazes (Brent's E-ZTax Form VBProject.vbp)
'A message from the commisioner (frmMessage)
'Brent Timothy Mergen
'24 March 2006
'This form gives some information about the IRS and tax returns.

Private Sub cmdReturnFrontpage_Click()
    frmMessage.Hide 'hides previous form
    frmFrontpage.Show 'brings you to a new form
End Sub

