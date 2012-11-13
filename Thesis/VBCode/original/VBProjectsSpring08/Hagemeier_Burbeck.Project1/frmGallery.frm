VERSION 5.00
Begin VB.Form frmGallery 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gallery"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17385
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   17385
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   16875
      TabIndex        =   4
      Top             =   9960
      Width           =   16935
   End
   Begin VB.PictureBox picCountry 
      Height          =   8775
      Left            =   4440
      ScaleHeight     =   8715
      ScaleWidth      =   12675
      TabIndex        =   3
      Top             =   1080
      Width           =   12735
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "MainMenu"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox txtCountry 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdChangeCountry 
      Caption         =   "Change Country"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Image imgGlobe 
      Height          =   4650
      Left            =   -120
      Picture         =   "frmGallery.frx":0000
      Top             =   4920
      Width           =   4665
   End
   Begin VB.Label lblcountry 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter a country name and hit Change Country to vew a picture"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblPrompt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to the picture Gallery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   5
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmGallery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmGallery (Gallery)
'Author: Nate Burbeck
'Date Written: 30 March 2008
'Objective: lets the user see a picture for each country in our directory
Option Explicit

Private Sub cmdChangeCountry_Click()
    Dim Country As String
    picResults.Cls
  
    Country = App.Path & "\Images\" & txtCountry.Text & ".bmp"
    
    If Dir(Country) = "" Then
        MsgBox "Error!  There is no picture of that country in our directory!", , "Error!"
    Else
        picResults.Print "Showing: "; Country
        ' The LoadPicture function actually changes the picture in the picturebox
        picCountry.Picture = LoadPicture(Country)
    End If
End Sub

Private Sub cmdMainMenu_Click()
    frmGallery.Hide
    frmMainMenu.Show
End Sub
