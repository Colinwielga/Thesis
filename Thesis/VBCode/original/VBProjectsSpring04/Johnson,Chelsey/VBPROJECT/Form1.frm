VERSION 5.00
Begin VB.Form MapLondon 
   BackColor       =   &H00FF0000&
   Caption         =   "Map of London"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   18
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtinstruction 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   17
      Text            =   "Please click on one of the districts to view the top famous sites of that district."
      Top             =   2160
      Width           =   10095
   End
   Begin VB.TextBox txtmap 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1080
      TabIndex        =   16
      Text            =   "This is a Map of London broken up into the main 16 districts.  "
      Top             =   1320
      Width           =   12255
   End
   Begin VB.PictureBox picWoolwich 
      Height          =   855
      Left            =   5760
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   15
      Top             =   4200
      Width           =   855
   End
   Begin VB.PictureBox picGreenwich 
      Height          =   735
      Left            =   4920
      Picture         =   "Form1.frx":02E2
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   5400
      Width           =   855
   End
   Begin VB.PictureBox picIsle 
      Height          =   975
      Left            =   4920
      Picture         =   "Form1.frx":0688
      ScaleHeight     =   915
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   4560
      Width           =   855
   End
   Begin VB.PictureBox picWapping 
      Height          =   855
      Left            =   4080
      Picture         =   "Form1.frx":0A4E
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   4560
      Width           =   855
   End
   Begin VB.PictureBox picTower 
      Height          =   735
      Left            =   3480
      Picture         =   "Form1.frx":0DAD
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   11
      Top             =   4320
      Width           =   735
   End
   Begin VB.PictureBox picCity2 
      Height          =   615
      Left            =   3600
      Picture         =   "Form1.frx":10B7
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   3720
      Width           =   975
   End
   Begin VB.PictureBox picCity 
      Height          =   855
      Left            =   3120
      Picture         =   "Form1.frx":1319
      ScaleHeight     =   795
      ScaleWidth      =   675
      TabIndex        =   9
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox picWestminister 
      Height          =   975
      Left            =   2400
      Picture         =   "Form1.frx":16A2
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.PictureBox picWestEnd 
      Height          =   495
      Left            =   2400
      Picture         =   "Form1.frx":1C4C
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.PictureBox picRegent 
      Height          =   735
      Left            =   1560
      Picture         =   "Form1.frx":1F8F
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.PictureBox picJames 
      Height          =   495
      Left            =   1560
      Picture         =   "Form1.frx":238F
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   1560
      Picture         =   "Form1.frx":266D
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   4920
      Width           =   855
   End
   Begin VB.PictureBox picPilmico 
      Height          =   735
      Left            =   1800
      Picture         =   "Form1.frx":28D3
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   5280
      Width           =   855
   End
   Begin VB.PictureBox picKnightsbridge 
      Height          =   855
      Left            =   360
      Picture         =   "Form1.frx":2C06
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.PictureBox picBattersea 
      Height          =   735
      Left            =   960
      Picture         =   "Form1.frx":3053
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   5280
      Width           =   855
   End
   Begin VB.PictureBox picChelsea 
      Height          =   615
      Left            =   240
      Picture         =   "Form1.frx":33B2
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Welcome To:  Discovering London"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   20
      Top             =   240
      Width           =   9855
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   11880
      Width           =   2775
   End
End
Attribute VB_Name = "MapLondon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: MapLondon (Form1.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of the Project: The purpose of this project is for the user to be able to learn and discover many different sites
                        'within London.  This is accomplished by breaking up London into 16 different catagories.
'Purpose of the Form: The purpose of this form is show the user the 16 different areas of London and have them choose where
                        'they would like to go from there.
'Option Explicit is a command used to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit

Private Sub cmdquit_Click()
End
End Sub

Private Sub picBattersea_Click()
'The form for Battersea is now showing
MapLondon.Hide
Battersea.Show
End Sub

Private Sub picChelsea_Click()
'The form for Chelsea is now showing
MapLondon.Hide
Chelsea.Show
End Sub

Private Sub picCity_Click()
'The form for City is now showing
MapLondon.Hide
City.Show
End Sub

Private Sub picCity2_Click()
'The form for City 2 is now showing
MapLondon.Hide
City2.Show
End Sub

Private Sub picGreenwich_Click()
'The form for Greenwich is now showing
MapLondon.Hide
Greenwich.Show
End Sub

Private Sub picIsle_Click()
'The form for Isle of Dogs is now showing
MapLondon.Hide
IsleDogs.Show
End Sub

Private Sub picJames_Click()
'The form for St. James is now showing
MapLondon.Hide
StJames.Show
End Sub

Private Sub picKnightsbridge_Click()
'The form for Knightsbridge is now showing
MapLondon.Hide
Knightsbridge.Show

End Sub

Private Sub picPilmico_Click()
'The form for Pilmico is now showing
MapLondon.Hide
Pimlico.Show
End Sub

Private Sub picRegent_Click()
'The form for Regent Street is now showing
MapLondon.Hide
RegentStreet.Show
End Sub

Private Sub picTower_Click()
'The form for Tower is now showing
MapLondon.Hide
Tower.Show
End Sub

Private Sub Picture1_Click()
'The form for Victoria is now showing
MapLondon.Hide
Victoria.Show
End Sub

Private Sub picWapping_Click()
'The form for Wapping is now showing
MapLondon.Hide
Wapping.Show
End Sub

Private Sub picWestEnd_Click()
'The form for West End is now showing
MapLondon.Hide
WestEnd.Show
End Sub

Private Sub picWestminister_Click()
'The form for Westminister is now showing
MapLondon.Hide
Westminister.Show
End Sub

Private Sub picWoolwich_Click()
'The form for Woolwich is now showing
MapLondon.Hide
Woolwich.Show
End Sub
