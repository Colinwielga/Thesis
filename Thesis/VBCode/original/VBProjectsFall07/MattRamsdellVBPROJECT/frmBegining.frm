VERSION 5.00
Begin VB.Form frmBegining 
   BackColor       =   &H00FF0000&
   Caption         =   "Ducks"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWisconsin 
      Caption         =   "Wisconsin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdIowa 
      Caption         =   "Iowa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdMinnesota 
      Caption         =   "Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2760
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdSouthDakota 
      Caption         =   "South Dakota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdNorthDakota 
      Caption         =   "North Dakota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the state below that you will be hunting or are hunting in."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1095
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label lblBegining 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBegining.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1575
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmBegining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdIowa_Click()
'Takes user to the Iowa page
frmBegining.Hide
frmIowa.Show
End Sub

Private Sub cmdMinnesota_Click()
'Takes User to the Minnesota page
frmBegining.Hide
frmMinnesota.Show
End Sub

Private Sub cmdNorthDakota_Click()
'Takes user to the North Dakota page
frmBegining.Hide
frmNorthDakota.Show
End Sub

Private Sub cmdSouthDakota_Click()
'Takes user to the South Dakota page
frmBegining.Hide
frmSouthDakota.Show
End Sub

Private Sub cmdWisconsin_Click()
'Takes the user to the Wisconsin page
frmBegining.Hide
frmWisconsin.Show
End Sub

