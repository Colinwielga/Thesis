VERSION 5.00
Begin VB.Form frmGeneral 
   BackColor       =   &H8000000B&
   Caption         =   "General Page"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   Picture         =   "FormJetLi.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCitation 
      Caption         =   "Works Cited"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main Page"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CommandButton cmdSale 
      Caption         =   "Movies Available for Sale"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton cmdDescriptions 
      Caption         =   "Movie Listings"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdBio 
      Caption         =   "Short Biography"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lblFong 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fong Sai-Yuk"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      TabIndex        =   5
      Top             =   1080
      Width           =   4815
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Planet of Jet Li
'Form Name: frmGeneral
'Author: Chakong Thao
'Date Written: Sunday, Oct. 29th
'Form Objective: This form provides a general page where everything
                'the program offers are categorized into command buttons.
                
Option Explicit

Private Sub cmdBio_Click()  'This button brings user to the page containing the biography
    frmGeneral.Hide
    frmBio.Show
End Sub

Private Sub cmdCitation_Click() 'This brings user to citation page, showing which websites were used for this program
    frmGeneral.Hide
    frmCitations.Show
    
    Counter = 0
    Open App.Path & "\Citations.txt" For Input As #1
    
    Do Until EOF(1)
        Input #1, Cite
        Counter = Counter + 1
        Citations(Counter) = Cite
    Loop
    
    Close #1
    
End Sub

Private Sub cmdDescriptions_Click() 'This button brings user to the page with movie listings
    frmGeneral.Hide
    frmMovies.Show
End Sub

Private Sub cmdMain_Click() 'This brings user back to beginning page
    frmGeneral.Hide
    frmJetLi.Show
End Sub

Private Sub cmdSale_Click() 'This brings user to the page with the movie search textbox
    frmGeneral.Hide
    frmMovieSale.Show
End Sub
