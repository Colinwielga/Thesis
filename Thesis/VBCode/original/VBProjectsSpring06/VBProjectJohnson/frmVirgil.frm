VERSION 5.00
Begin VB.Form frmVirgil 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virgil Michel- Additional Information"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInfoVirgil 
      BackColor       =   &H000000FF&
      Caption         =   "View General Information For Virgil Michel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdFloorVirgil 
      BackColor       =   &H000000FF&
      Caption         =   "View Floor Plan For Virgil Michel By Clicking On the Icon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back To Draft"
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
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
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
      Left            =   6360
      TabIndex        =   5
      Top             =   4920
      Width           =   2055
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H00FF0000&
      Class           =   "AcroExch.Document.7"
      DisplayType     =   1  'Icon
      Height          =   975
      Left            =   360
      OleObjectBlob   =   "frmVirgil.frx":0000
      SourceDoc       =   "\\ad\homedir$\Students\KMJOHNSON\My Documents\VirgilMichel2nd.pdf"
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblVirgil 
      BackColor       =   &H000000FF&
      Caption         =   $"frmVirgil.frx":4AA18
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
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "frmVirgil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. Johns Housing Project
' Virgil Form
' Written By Kyle Johnson
' 3/22/06
' this form displays additional information about Virgil house including
' a brief description of the house,  and also a floor plan of the house


Private Sub cmdBack_Click()
    'navigates from virgil page to the options page
    frmOptions.Visible = True
    frmVirgil.Visible = False
    
    lblVirgil.Visible = False
    
End Sub
    

    
Private Sub cmdInfoVirgil_Click()
    'display the additional information for virgil michel
    lblVirgil.Visible = True
    
End Sub
    
Private Sub cmdPicVirgil_Click()
    'navigates from the virgil page to the additional pictures of virgil
    frmPicVirgil.Visible = True
    frmVirgil.Visible = False
    
End Sub
    
Private Sub Form_Load()
    'makes the label initially hidden when form is loaded
        lblVirgil.Visible = False

End Sub
