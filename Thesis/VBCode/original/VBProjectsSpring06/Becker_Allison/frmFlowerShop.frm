VERSION 5.00
Begin VB.Form frmFlowerShop 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Flower Shop"
   ClientHeight    =   9315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   48
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFlowerShop.frx":0000
   ScaleHeight     =   9315
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSpecial 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click to View March Special"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   6720
      Picture         =   "frmFlowerShop.frx":AFA2
      ScaleHeight     =   3555
      ScaleWidth      =   3555
      TabIndex        =   7
      Top             =   5280
      Width           =   3615
   End
   Begin VB.CommandButton cmdLocations 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click to View Locations throughout Minnesota"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFF00&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdFlowers 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click to View Flowers Available for Purchase"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click to Learn about Flowers For U!  "
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click to View Prices for Flowers "
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Designed By Allison Becker"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   8400
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFlowerShop.frx":35014
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1455
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   10455
   End
   Begin VB.Label lblFlowers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Flowers For U!"
      ForeColor       =   &H00FFFF00&
      Height          =   2055
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
End
Attribute VB_Name = "frmFlowerShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Flowers For U! (FlowerShop.vbp)
'Form Name: (frmFlowerShop)
'Author: Allison Becker
'Date Written: 3/23/06
'Objective: The objective of this page is giving the user options on what they would
'like to view. It is the navigation page to all other pages that can be seen
'throughout the whole Flowers For U! program. There is also a quit button located
'on this page which stops the program all together.
Option Explicit

Private Sub cmdAbout_Click()
    frmFlowerShop.Hide
    frmAboutFlowers.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdFlowers_Click()
    frmFlowersAvailable.Show
    frmFlowerShop.Hide
End Sub

Private Sub cmdOrder_Click()
    frmAboutFlowers.Show
    frmFlowerShop.Hide
End Sub

Private Sub cmdLocations_Click()
    frmLocations.Show
    frmFlowerShop.Hide
End Sub

Private Sub cmdPrice_Click()
    frmFlowerShop.Hide
    frmPrices.Show
End Sub


Private Sub cmdSpecial_Click()
    frmFlowerShop.Hide
    frmSpecial.Show
End Sub


