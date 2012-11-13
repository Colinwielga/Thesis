VERSION 5.00
Begin VB.Form frmCitations 
   BackColor       =   &H80000007&
   Caption         =   "Works Cited"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmCitations.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Citations"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   3
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   8160
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   3840
      ScaleHeight     =   5235
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   5160
      Width           =   9015
   End
   Begin VB.Label lblCite 
      BackStyle       =   0  'Transparent
      Caption         =   "  Works Cited"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
   End
End
Attribute VB_Name = "frmCitations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Planet of Jet Li
'Form Name: frmCitations
'Author: Chakong Thao
'Date Written: Thursday, Nov. 2nd
'Form Objective: The things displayed on this form steps away from
                'information about Jet Li and his movies for sale.
                'Instead, it lists all the websites used to help
                'design this project as far as the biography to
                'images on each interface.
                
Private Sub cmdBack_Click() 'This brings user back to General page
    frmCitations.Hide
    frmGeneral.Show
End Sub

Private Sub cmdMain_Click() 'This brings user back to beginning page
    frmCitations.Hide
    frmJetLi.Show
End Sub

Private Sub cmdShow_Click()
    picResults.Cls
    For Pos = 1 To Counter
        picResults.Print Citations(Pos)
    Next Pos
End Sub
