VERSION 5.00
Begin VB.Form frmParisAttractions 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080C0FF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add Attractions To Budget"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CheckBox chkLuve 
      Caption         =   "Check3"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   5880
      Width           =   255
   End
   Begin VB.CheckBox chkInvalides 
      Caption         =   "Check2"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chkEiffel 
      Caption         =   "Check2"
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox chkMoulin 
      Caption         =   "Check2"
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   6120
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   3840
      Picture         =   "frmParisAttractions.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   5160
      Width           =   2895
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2880
         TabIndex        =   5
         Top             =   600
         Width           =   75
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   2175
      Left            =   3120
      Picture         =   "frmParisAttractions.frx":179C2
      ScaleHeight     =   2115
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   480
      Picture         =   "frmParisAttractions.frx":2A5A4
      ScaleHeight     =   2235
      ScaleWidth      =   2835
      TabIndex        =   1
      Top             =   3360
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   6720
      Picture         =   "frmParisAttractions.frx":3F766
      ScaleHeight     =   4395
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H0080FFFF&
      Caption         =   "Attractions In Paris"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   13
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label lblLuve 
      BackColor       =   &H0080FFFF&
      Caption         =   "Visit The Luve $11"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label lblInvalides 
      BackColor       =   &H0080FFFF&
      Caption         =   "Tour Esplande Des Invalides $6.50"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblEiffel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ride to the Top of the Eiffel Tower $9"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   7
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label lblMoulin 
      BackColor       =   &H0080FFFF&
      Caption         =   "Moulin Rouge: Dinner and a Show $130"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      TabIndex        =   4
      Top             =   6000
      Width           =   1935
   End
End
Attribute VB_Name = "frmParisAttractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmParisAttractions
'Jessica Florek
'Written: 3/8/09
'Objective: This form has a variety of options for entertainment in Paris that the user can choose
'and will be added to their budgets.


Option Explicit

Private Sub cmdAdd_Click()
'depending on the check boxes that are checked, the corresponding prices are subtracted from the budget and added to parisattractioncost that is used later for the budget summary
If chkInvalides Then
    budget = budget - 6.5
    parisattractioncost = parisattractioncost + 6.5
End If
If chkEiffel Then
    budget = budget - 9
    parisattractioncost = parisattractioncost + 9
End If
If chkLuve Then
    budget = budget - 11
    parisattractioncost = parisattractioncost + 11
End If
If chkMoulin Then
    budget = budget - 130
    parisattractioncost = parisattractioncost + 130
End If

paris = True
'makes the city information appear on the budget summary
frmParisAttractions.Hide
frmParis.Show

    
End Sub

Private Sub cmdBack_Click()
frmParisAttractions.Hide
frmParis.Show
End Sub
