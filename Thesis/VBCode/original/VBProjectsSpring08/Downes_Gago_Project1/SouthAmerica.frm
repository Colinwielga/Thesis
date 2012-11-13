VERSION 5.00
Begin VB.Form SouthAmerica 
   BackColor       =   &H00C00000&
   Caption         =   "South America"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form2"
   ScaleHeight     =   9000
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   5415
   End
   Begin VB.CommandButton cmdTagsNFlags 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TEST YOUR KNOWLEDGE OF THE FLAGS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   5535
   End
   Begin VB.CommandButton cmdTravel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TRAVEL"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Text            =   "                              What Would You Like To Do?"
      Top             =   120
      Width           =   12855
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   8160
      Left            =   5760
      Picture         =   "SouthAmerica.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   6960
   End
End
Attribute VB_Name = "SouthAmerica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: SouthAmerica.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  This Form gives the user the ability to switch between forms

Option Explicit
'Hides the South America Form and Shows the Main Form
Private Sub cmdBack_Click()
SouthAmerica.Hide
Main.Show
End Sub
'Hides the SouthAmerica Form and the TagsNFlags Form Shows
Private Sub cmdTagsNFlags_Click()
SouthAmerica.Hide
TagsNFlags.Show

End Sub
'Hides the SouthAmerica Form and Shows the Travel Form
Private Sub cmdTravel_Click()
SouthAmerica.Hide
Travel.Show
End Sub
