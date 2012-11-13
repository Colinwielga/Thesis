VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H008080FF&
   Caption         =   "Cruise Home Page"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   4095
      Left            =   6120
      Picture         =   "HomePage.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   5715
      TabIndex        =   5
      Top             =   1440
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   -360
      Picture         =   "HomePage.frx":AA23
      ScaleHeight     =   4035
      ScaleWidth      =   6075
      TabIndex        =   4
      Top             =   1440
      Width           =   6135
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton cmdAlaska 
      BackColor       =   &H0080FFFF&
      Caption         =   "Alaskan Cruise"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton cmdCaribbean 
      BackColor       =   &H0080FFFF&
      Caption         =   "Caribbean Cruise"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label lblCruise 
      BackColor       =   &H0080FFFF&
      Caption         =   $"HomePage.frx":1DBFD
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmHome
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/14/2009
'Objective: This is our startup form where the user can choose to view either an Alaskan
'Cruise or a Caribbean Cruise.
'The overall objective for our project is for the user to view information regarding one
'of two types of cruises, and the information they can view includes: type of cruise,
'activities on the ship (toddler, teen, adult), dining options and food menus, port
'destinations, room sizes and prices, flight information, booking information, and creating
'an ID card for when the user boards the cruise ship.

Private Sub cmdAlaska_Click()
frmHome.Hide
frmAlaskanHome.Show
End Sub

Private Sub cmdCaribbean_Click()
frmHome.Hide
frmCaribbeanHome.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

