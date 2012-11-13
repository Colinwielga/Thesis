VERSION 5.00
Begin VB.Form frm2 
   BackColor       =   &H00FF8080&
   Caption         =   "Ideal Body Weight"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form2"
   ScaleHeight     =   5370
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdfemale 
      Caption         =   "Click me "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdfrm1 
      Caption         =   "Go to Food Pyramid"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton cmdmale 
      Caption         =   "Click me"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Healthy Living
'frm2
'Ben Morris
'March 21
'Homepage the ideal weight calcualtor


Private Sub cmdfemale_Click()
    frm2.Hide
    frm10.Show
    'hides the idela weight form and opens the form for calculating a females ideal weight
End Sub

Private Sub cmdfrm1_Click()
    frm1.Show
    frm2.Hide
    'shows the ideal weight form and hides the pyramid form
End Sub

Private Sub cmdmale_Click()
    frm2.Hide
    frm9.Show
    'hides the idela weight form and opens the form for calculating a males ideal weight
End Sub
