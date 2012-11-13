VERSION 5.00
Begin VB.Form frmSources 
   BackColor       =   &H00800000&
   Caption         =   "Reference Sources"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
      Caption         =   "www.sports-db.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
      Caption         =   "www.myhealth.reidhosp.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      Caption         =   "www.building-muscle101.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "www.southernpowerliftingclub.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Images courtesy of:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "ACE (American Council on Exercise) Personal Trainer Manual: 3rd Edition"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Equations and definitions courtesy of:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FITNESS ASSESSOR
'frmSources
'Nick Schuster
'March 26, 2008

'This form shows the user the references used to create this program
Option Explicit
'To return to the previous form
Private Sub cmdBack_Click()
frmSources.Hide
frmWelcome.Show
End Sub
