VERSION 5.00
Begin VB.Form frmDirections 
   BackColor       =   &H0080FFFF&
   Caption         =   "Find Directions"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   Picture         =   "frm2.frx":0000
   ScaleHeight     =   9060
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUniversity6 
      Caption         =   "University Of California"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   12
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdUniversity5 
      Caption         =   "University Of Illinois"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdUniversity4 
      Caption         =   "Harvard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   5280
      Width           =   1935
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< Back to Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   8160
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   3120
      Picture         =   "frm2.frx":400F1
      ScaleHeight     =   7305
      ScaleWidth      =   9225
      TabIndex        =   3
      Top             =   600
      Width           =   9255
   End
   Begin VB.CommandButton cmdUniversity3 
      Caption         =   "St. John's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdUniversity2 
      Caption         =   "University Of Minesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdUniversity1 
      Caption         =   "St. Mary's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblCity 
      BackColor       =   &H0032D6D6&
      Caption         =   "City Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblState 
      BackColor       =   &H0032D6D6&
      Caption         =   "State Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmDirections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: College Bound
'Form Name: Directions
' Authors: Magdalena Adamczyk & Leszek Nowacki
' 9-25 March 2009
' This form is designed to let the user see maps showing the universities
' by choising the appropriate option and clicking the button with the name of the university the user can
' see the map of the state in which the university is located or the map showing the universitie's campus from city level


Private Sub cmdUniversity1_Click() 'this button loads the picture of St. Mary's University
picResults.Cls
If opt2 = True Then 'if this option button is marked the city level map will be displayed
    Picture1 = App.Path & "\image1.jpg"
    Set picResults.Picture = LoadPicture(Picture1)
End If

If opt1 = True Then 'if this option button is marked the state map will be displayed
    Picture11 = App.Path & "\image11.bmp"
    Set picResults.Picture = LoadPicture(Picture11)
End If

End Sub


Private Sub cmdUniversity2_Click() 'this button loads the picture of University of Minnesota
picResults.Cls

If opt2 = True Then 'if this option button is marked the city level map will be displayed
    Picture2 = App.Path & "\image2.jpg"
    Set picResults.Picture = LoadPicture(Picture2)
End If

If opt1 = True Then 'if this option button is marked the state map will be displayed
    Picture22 = App.Path & "\image22.bmp"
    Set picResults.Picture = LoadPicture(Picture22)
End If
End Sub

Private Sub cmdUniversity3_Click() 'this button loads the picture of St. John's University
picResults.Cls
If opt2 = True Then 'if this option button is marked the city level map will be displayed
    Picture3 = App.Path & "\image3.jpg"
    Set picResults.Picture = LoadPicture(Picture3)
End If

If opt1 = True Then 'if this option button is marked the state map will be displayed
    Picture33 = App.Path & "\image33.bmp"
    Set picResults.Picture = LoadPicture(Picture33)
End If
End Sub

Private Sub cmdUniversity4_Click() 'this button loads the picture of Harvard
picResults.Cls
If opt2 = True Then 'if this option button is marked the city level map will be displayed
    Picture4 = App.Path & "\image4.jpg"
    Set picResults.Picture = LoadPicture(Picture4)
End If

If opt1 = True Then 'if this option button is marked the state map will be displayed
    Picture44 = App.Path & "\image44.bmp"
    Set picResults.Picture = LoadPicture(Picture44)
End If
End Sub


Private Sub cmdUniversity5_Click() ' 'this button loads the picture of university of Illinois
picResults.Cls
If opt2 = True Then 'if this option button is marked the city level map will be displayed
    Picture5 = App.Path & "\image5.jpg"
    Set picResults.Picture = LoadPicture(Picture5)
End If

If opt1 = True Then 'if this option button is marked the state map will be displayed
    Picture55 = App.Path & "\image55.bmp"
    Set picResults.Picture = LoadPicture(Picture55)
End If
End Sub

Private Sub cmdUniversity6_Click() 'this button loads the picture of University of California
picResults.Cls
If opt2 = True Then 'if this option button is marked the city level map will be displayed
    Picture6 = App.Path & "\image6.jpg"
    Set picResults.Picture = LoadPicture(Picture6)
End If

If opt1 = True Then 'if this option button is marked the state map will be displayed
    Picture66 = App.Path & "\image66.bmp"
    Set picResults.Picture = LoadPicture(Picture66)
End If
End Sub
Private Sub cmdBack_Click() ' this button hides the current form and displays other form which lets the user search through a list of universities according to the tuition price and major
frmDirections.Hide
frmUniversitySearch.Show
End Sub

Private Sub cmdQuit_Click() 'this button ends the program
End
End Sub

