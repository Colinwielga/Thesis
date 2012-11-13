VERSION 5.00
Begin VB.Form frmMexicanVeggie 
   BackColor       =   &H00008000&
   Caption         =   "Mexican Vegetarian"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMexicanVeggie.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   8925
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMVQuit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      MaskColor       =   &H00004000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdMVReturn 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      MaskColor       =   &H00004000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Click Here to See What You Need for Chipotle Salsa"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7320
      MaskColor       =   &H00004000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H00008000&
      Caption         =   "Chipotle Salsa"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   3630
      Left            =   1680
      Picture         =   "frmMexicanVeggie.frx":08CA
      Top             =   1800
      Width           =   5475
   End
   Begin VB.Label lblMVStepOne 
      BackColor       =   &H00008000&
      Caption         =   "1.Combine all ingredients in a bowl and mix well."
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   6480
      Width           =   9375
   End
   Begin VB.Label lblMVSteps 
      BackColor       =   &H00008000&
      Caption         =   "There are ONLY ONE EASY step to make Mexican Chipotle Salsa:"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1335
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   5880
      Width           =   10215
   End
End
Attribute VB_Name = "frmMexicanVeggie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMVQuit_Click()
End
End Sub

Private Sub cmdMVReturn_Click()

'Return to Homepage
frmCountries.Show
frmMexicanVeggie.Hide

End Sub

Private Sub cmdNext_Click()
'See modules for more.
groceryfile = "\Recipes\mexicanVR.txt"

'Next Step
frmMexicanVeggie.Hide
frmGroceryStore.Show

End Sub
