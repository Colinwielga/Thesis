VERSION 5.00
Begin VB.Form China 
   Caption         =   "China"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3675
      ScaleWidth      =   6675
      TabIndex        =   3
      Top             =   4200
      Width           =   6735
   End
   Begin VB.CommandButton cmdQuit 
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
      Height          =   1215
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton cmdTrivia 
      BackColor       =   &H00C0E0FF&
      Caption         =   "TRIVIA"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdPhoto 
      BackColor       =   &H00C0E0FF&
      Caption         =   "PICTURES"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "China"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: China.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  The objective of this form is to display useful information
'on China through pictures and message boxes

Option Explicit
'The China Form Disappears while the ChinaPhoto Form appears
Private Sub cmdPhoto_Click()
China.Hide
ChinaPhoto.Show
China.Picture = LoadPicture(App.Path & "\China\Chinaflag.jpg")  'The Chinese flag is displayed in the background after the image is loaded
End Sub
'The China Form Disappears while the Asia Form Appears
Private Sub cmdQuit_Click()
China.Hide
Asia.Show
End Sub
'This sub command displays useful information on china
Private Sub cmdTrivia_Click()

picResults.Cls      'Clears the picture box

picResults.Enabled = True       'Causes the next lines to be printed after the button is clicked
picResults.Print "In October 1964, the People's Republic of China exploded her first atomic bomb."
picResults.Print "On October 27, 1966, China exploded her first nuclear bomb from a guided missile."
picResults.Print "On June 17, 1967 China exploded her first hydrogen bomb."

picResults.Print ""
picResults.Print ""
picResults.Print ""
picResults.Print ".cn is the internet code for Chinese websites."
picResults.Print "By the year 2003, China had 3 internet service providers and about 45.8 million "
picResults.Print "internet users."

picResults.Print ""
picResults.Print ""
picResults.Print ""
picResults.Print "Turkey's area is 780,580 square kilometres."
picResults.Print "Mozambique's area is 801,590 square kilometres."
picResults.Print "Australia's area is 7,686,850 square kilometres."
picResults.Print "When you add these three areas, you get 9,269,020 square kilometres,"
picResults.Print "and China's area is 9,596,960 square kilometres!"

picResults.Print ""
picResults.Print ""
picResults.Print ""
picResults.Print ""

End Sub
