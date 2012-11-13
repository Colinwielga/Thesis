VERSION 5.00
Begin VB.Form frmAmerican 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amiercan"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAmerican.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6060
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdClickHere 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Click Here to See What You NEED for Beef Burgers"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6960
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label lblDishName 
      BackColor       =   &H00004080&
      Caption         =   "Beef Burger!"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   855
      Left            =   6240
      TabIndex        =   6
      Top             =   840
      Width           =   4455
   End
   Begin VB.Image imgAmerican 
      Height          =   6090
      Left            =   -120
      Picture         =   "frmAmerican.frx":08CA
      Top             =   0
      Width           =   6090
   End
   Begin VB.Label lblAStepTwo 
      BackColor       =   &H00004080&
      Caption         =   "2. Broil in oven, or grill"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   2
      Left            =   6120
      TabIndex        =   2
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label lblAStepOne 
      BackColor       =   &H00004080&
      Caption         =   "1. Mix all ingredients together"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Index           =   1
      Left            =   6120
      TabIndex        =   1
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label lblStepsA 
      BackColor       =   &H00004080&
      Caption         =   "There are only TWO simple steps to make ..."
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Index           =   0
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmAmerican"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAQuit_Click()

End

End Sub

Private Sub cmdClickHere_Click()

groceryfile = "\Recipes\americanR.txt "

'Next Step
frmAmerican.Hide
frmGroceryStore.Show





End Sub

Private Sub cmdReturn_Click()

'Return to Homepage
frmCountries.Show
frmAmerican.Hide

End Sub

