VERSION 5.00
Begin VB.Form Asia 
   BackColor       =   &H000000C0&
   Caption         =   "Asia"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Text            =   "                    Intrested in Asia?    Learn More About Desired Country"
      Top             =   120
      Width           =   10695
   End
   Begin VB.CommandButton CmdChina 
      BackColor       =   &H80000016&
      Caption         =   "CHINA"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   3975
   End
   Begin VB.CommandButton cmdRussia 
      BackColor       =   &H80000016&
      Caption         =   "RUSSIA"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   3975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000016&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   4200
      Picture         =   "Asia.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   6360
   End
End
Attribute VB_Name = "Asia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: Asia.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  to give the user the ability to switch forms

'The Asia Form disapears and brings the user back to the main form
Private Sub cmdBack_Click()
Asia.Hide
Main.Show
End Sub
'Hides the Asia Form, and then Shows the China Form
Private Sub CmdChina_Click()
Asia.Hide
China.Show
China.Picture = LoadPicture(App.Path & "\China\Chinaflag.jpg")  'The form background becomes the loaded picture here
End Sub
'Hides the Asia Form and Shows the Russia Form
Private Sub cmdRussia_Click()
Asia.Hide
Russia.Show
Russia.Picture = LoadPicture(App.Path & "\moscow3.jpg") 'The form background becomes the loaded picture here
End Sub
