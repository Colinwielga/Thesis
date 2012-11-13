VERSION 5.00
Begin VB.Form frmdefpositions 
   BackColor       =   &H8000000D&
   Caption         =   "Defensive Players"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Back to Main Menu"
      Height          =   975
      Left            =   4680
      TabIndex        =   0
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   5640
      Picture         =   "frmdefpositions.frx":0000
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lbldbs 
      Caption         =   "Defensive Backs"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Defensive End"
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lbldt 
      Caption         =   "Defensive Tackle"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Image imgdt 
      Height          =   3000
      Left            =   3120
      Picture         =   "frmdefpositions.frx":154E2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "DEFENSIVE PLAYERS"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   4935
   End
   Begin VB.Image imgde 
      Height          =   2175
      Left            =   8040
      Picture         =   "frmdefpositions.frx":2A9C4
      Top             =   3000
      Width           =   3075
   End
   Begin VB.Image imglbs 
      Height          =   3000
      Left            =   360
      Picture         =   "frmdefpositions.frx":406EE
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lbllbs 
      Caption         =   "Linebackers"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
End
Attribute VB_Name = "frmdefpositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'2006 NFL Draft Simulator (Draft.vbp)
'Draft Busts? or MVP's?(frmNFLDraft)
'Andy Lyons
'March 24, 2006
'This form is used to look at the Defensive Players that are eligible for the 2006 NFL Draft. By clicking on each individual picture, it allows the user to read their personal profile.
'returns user to main menu
Private Sub cmdreturn_Click()
    frmNFLDraft.Show
    frmdefpositions.Hide
End Sub

Private Sub Image1_Click()
    frmdbs.Show
    frmdefpositions.Hide
End Sub

Private Sub imgde_Click()
    frmends.Show
    frmdefpositions.Hide
End Sub

Private Sub imgdt_Click()
    frmsdts.Show
    frmdefpositions.Hide
End Sub

Private Sub imglbs_Click()
    frmlbs.Show
    frmdefpositions.Hide
End Sub
