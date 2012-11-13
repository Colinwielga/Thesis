VERSION 5.00
Begin VB.Form Dessert 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form4"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form4"
   ScaleHeight     =   5130
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitch3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Back to Main Form"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton cmdAsian3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Dessert 2: Asian"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdItalian3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Dessert 1: Italian"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label lblDessert 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Dessert"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Dessert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAsian3_Click()
Dessert.Hide 'this hides the dessert form
AsianDessert.Show   'this shows the asian dessert form
End Sub

Private Sub cmdItalian3_Click()
Dessert.Hide    'this hides the dessert form
ItalianDessert.Show 'this shows the italian dessert form
End Sub

Private Sub cmdSwitch3_Click()
Dessert.Hide    'this hides the dessert form
Main.Show       'this shows the main form
End Sub
