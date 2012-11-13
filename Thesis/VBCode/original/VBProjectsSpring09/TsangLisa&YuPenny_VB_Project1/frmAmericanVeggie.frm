VERSION 5.00
Begin VB.Form frmAmericanVeggie 
   BackColor       =   &H00C0FFC0&
   Caption         =   "American Vegetarian"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmAmericanVeggie.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmAmericanVeggie.frx":08CA
   ScaleHeight     =   8955
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAVQuit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturnAV 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmdClickHereAV 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Click Here to see what you Need for Caesar Salad"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H8000000E&
      Caption         =   "Caesar Salad "
      BeginProperty Font 
         Name            =   "Script"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   6180
   End
   Begin VB.Image imgAmericanV 
      Height          =   3660
      Left            =   960
      Picture         =   "frmAmericanVeggie.frx":2D990C
      Top             =   960
      Width           =   6180
   End
   Begin VB.Label lblStepsAVThree 
      BackColor       =   &H00C0FFC0&
      Caption         =   $"frmAmericanVeggie.frx":32335E
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   7800
      Width           =   8415
   End
   Begin VB.Label lblStepsAVOne 
      BackColor       =   &H00C0FFC0&
      Caption         =   "1.Wash romaine; remove outside leaves. Break tender leaves crosswise into pieces about 1 inch wide"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   5400
      Width           =   8415
   End
   Begin VB.Label lblStepsAVTwo 
      BackColor       =   &H00C0FFC0&
      Caption         =   $"frmAmericanVeggie.frx":3233F1
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   6240
      Width           =   8415
   End
   Begin VB.Label lblStepsAV 
      BackColor       =   &H00C0FFC0&
      Caption         =   "There are Three steps to make Caesar Salad:"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   4800
      Width           =   8415
   End
End
Attribute VB_Name = "frmAmericanVeggie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAVQuit_Click()

End

End Sub

Private Sub cmdClickHereAV_Click()

groceryfile = "\Recipes\americanVR.txt "

'Next Step
frmAmericanVeggie.Hide
frmGroceryStore.Show

End Sub

Private Sub cmdReturnAV_Click()

'Return to Homepage
frmCountries.Show
frmAmericanVeggie.Hide

End Sub
