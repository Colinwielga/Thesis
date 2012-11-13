VERSION 5.00
Begin VB.Form frmMaur 
   BackColor       =   &H00FF0000&
   Caption         =   "Maur House- Additional Information"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H000000FF&
      Caption         =   "View Floor Plan For Maur By Clicking on The Icon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back To Draft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdInformation 
      BackColor       =   &H000000FF&
      Caption         =   "View General Information About Maur House"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00FF0000&
      Caption         =   "Project Created By:                 Kyle Johnson"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      TabIndex        =   5
      Top             =   5520
      Width           =   2055
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H00FF0000&
      Class           =   "AcroExch.Document.7"
      DisplayType     =   1  'Icon
      Height          =   1215
      Left            =   840
      OleObjectBlob   =   "frmMaur.frx":0000
      SourceDoc       =   "\\ad\homedir$\Students\KMJOHNSON\My Documents\Maur1st.pdf"
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblMaur 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMaur.frx":68C18
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   4440
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
   End
End
Attribute VB_Name = "frmMaur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. Johns Housing Project
' Maur Form
' Written By Kyle Johnson
' 3/22/06
' this form displays additional information about Maur house including
' a brief description of the house,  and also a floor plan of the house




Private Sub cmdBack_Click()
    ' takes user from maur form the the options form
    frmOptions.Visible = True
    frmMaur.Visible = False
    
    lblMaur.Visible = False
    End Sub
    

    
    Private Sub cmdInformation_Click()
    'displays the additional information for maur house
    lblMaur.Visible = True


End Sub

Private Sub Form_Load()
    'makes label initially hidden
    lblMaur.Visible = False

End Sub

