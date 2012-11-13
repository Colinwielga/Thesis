VERSION 5.00
Begin VB.Form Options 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form3"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   LinkTopic       =   "Form3"
   ScaleHeight     =   7545
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForm 
      Caption         =   "Go Back"
      Height          =   615
      Left            =   2280
      TabIndex        =   14
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox Picture6 
      Height          =   855
      Left            =   3960
      Picture         =   "Options.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   11
      Top             =   6360
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      Height          =   1095
      Left            =   3960
      Picture         =   "Options.frx":4A54
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      Height          =   735
      Left            =   3960
      Picture         =   "Options.frx":564B
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   2295
      Left            =   240
      Picture         =   "Options.frx":5EA8
      ScaleHeight     =   2235
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFC0C0&
      Height          =   2295
      Left            =   3480
      Picture         =   "Options.frx":A2A2
      ScaleHeight     =   2235
      ScaleWidth      =   2835
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFC0C0&
      Height          =   2055
      Left            =   240
      Picture         =   "Options.frx":D3B1
      ScaleHeight     =   1995
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "David Schonhardt"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "3: Western Red Cedar"
      Height          =   735
      Left            =   5280
      TabIndex        =   13
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "2: TimberTech Plastic"
      Height          =   735
      Left            =   5280
      TabIndex        =   12
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "2: Trex Composite"
      Height          =   615
      Left            =   5280
      TabIndex        =   9
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Types of Decking:"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Options for the Deck"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stair Landing"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Angled Corners (This platform has 3)"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stairs"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   2775
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Options (Options.frm), which is for the display of important options for the deck such as stairs, the type of material, and the inclusion of a landing.

Private Sub cmdForm_Click()
MainForm.Show
Options.Hide
End Sub
