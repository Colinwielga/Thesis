VERSION 5.00
Begin VB.Form Appetizer 
   BackColor       =   &H00C0C000&
   Caption         =   "Form2"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form2"
   ScaleHeight     =   4995
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H00FF8080&
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
      Height          =   855
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdAsian 
      BackColor       =   &H00FF8080&
      Caption         =   "Appetizer 2: Asian"
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
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdItalian 
      BackColor       =   &H00FF8080&
      Caption         =   "Appetizer 1: Italian"
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
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblAppetizer 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Appetizer"
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
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Appetizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAsian_Click()
AsianApp.Show       'this shows the asian appetizer form
Appetizer.Hide      'this hides the appetizer form
End Sub

Private Sub cmdItalian_Click()
ItalianApp.Show     'this shows the italian appetizer form
Appetizer.Hide      'this hides the appetizer form
End Sub

Private Sub cmdSwitch_Click()
Main.Show           'this shows the main form
Appetizer.Hide      'this hides the appetizer form
End Sub
