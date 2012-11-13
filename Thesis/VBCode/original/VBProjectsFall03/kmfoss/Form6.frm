VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FF8080&
   Caption         =   "Form6"
   ClientHeight    =   7995
   ClientLeft      =   510
   ClientTop       =   885
   ClientWidth     =   12225
   LinkTopic       =   "Form6"
   ScaleHeight     =   7995
   ScaleWidth      =   12225
   Begin VB.CommandButton cmdForm2 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   9
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form6.frx":0000
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   2535
      Left            =   6480
      TabIndex        =   8
      Top             =   4320
      Width           =   5295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form6.frx":0104
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   5655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form6.frx":0225
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   2535
      Left            =   6480
      TabIndex        =   6
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form6.frx":0347
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "RA Recruitment /Selection Committee"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   6255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "RA/CA Advisory Committee"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   6480
      TabIndex        =   3
      Top             =   3840
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Resource Room Committee"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   6480
      TabIndex        =   2
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Residential Life Office"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Collateral Assignments"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   3480
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForm2_Click()
Form6.Hide
Form2.Show
End Sub
