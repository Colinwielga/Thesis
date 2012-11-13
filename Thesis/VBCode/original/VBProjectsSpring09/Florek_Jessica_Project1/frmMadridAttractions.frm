VERSION 5.00
Begin VB.Form frmMadridAttractions 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form1"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check5 
      Caption         =   "Check5"
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   5520
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   7920
      TabIndex        =   13
      Top             =   9480
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   9240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C000C0&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C000C0&
      Caption         =   "Add Attractions to Budget"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      Height          =   2535
      Left            =   5400
      Picture         =   "frmMadridAttractions.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   3555
      TabIndex        =   4
      Top             =   600
      Width           =   3615
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   5880
      Picture         =   "frmMadridAttractions.frx":1D0E2
      ScaleHeight     =   1995
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   7440
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Height          =   3135
      Left            =   360
      Picture         =   "frmMadridAttractions.frx":3123C
      ScaleHeight     =   3075
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   6000
      Width           =   4935
   End
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   240
      Picture         =   "frmMadridAttractions.frx":85FCE
      ScaleHeight     =   2955
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   1560
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   6720
      Picture         =   "frmMadridAttractions.frx":AD418
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   3960
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Adventuros appetites Tapas Tour (visit local restaurants and eat a variety of local foods) $30 "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   6360
      TabIndex        =   16
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Puerto Del Sol FREE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   855
      Left            =   5400
      TabIndex        =   14
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "See a show at Teatro Royal $76"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   4680
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Attractions in Madrid"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Plaza Mayor FREE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   9240
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Museo De Prado $9"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   9480
      Width           =   2295
   End
End
Attribute VB_Name = "frmMadridAttractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmMadridAttractions
'Jessica Florek
'Written: 3/10/09
'Objective: This form has a variety of options for entertainment in Madrid that the user can choose
'and will be added to their budgets.


Option Explicit

'depending on the checkboxes that are checked, the corresponding activities will be calculated into the budget and added to madridattractionscost which will be used later during the budget summary
Private Sub cmdAdd_Click()
If Check1 Then
    budget = budget - 76
    madridattractioncost = madridattractioncost + 76
End If
If Check5 Then
    budget = budget - 30
    madridattractioncost = madridattractioncost + 30
End If
If Check4 Or Check2 Then
    budget = budget - 0
    madridattractioncost = madridattractioncost + 0
End If
If Check3 Then
    budget = budget - 9
    madridattractioncost = madridattractioncost + 9
End If

madrid = True
'this is used later during the budget summary, if a city is false it will not be displayed

frmMadridAttractions.Hide
frmMadrid.Show
End Sub

Private Sub cmdBack_Click()
frmMadridAttractions.Hide
frmMadrid.Show
End Sub
