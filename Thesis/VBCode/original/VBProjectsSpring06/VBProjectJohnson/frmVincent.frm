VERSION 5.00
Begin VB.Form frmVincent 
   BackColor       =   &H00FF0000&
   Caption         =   "Vincent- Additional Information"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
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
      Height          =   735
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdFloorVincent 
      BackColor       =   &H000000FF&
      Caption         =   "View Floor Plan For Vincent By Clicking On the Icon"
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
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdInfoVincent 
      BackColor       =   &H000000FF&
      Caption         =   "View General Information For Vincent"
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
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2055
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
      Top             =   5880
      Width           =   2055
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H00FF0000&
      Class           =   "AcroExch.Document.7"
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   600
      OleObjectBlob   =   "frmVincent.frx":0000
      SourceDoc       =   "\\ad\homedir$\Students\KMJOHNSON\My Documents\VincentCourtTypical2ndFloor.pdf"
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblVincent 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmVincent.frx":37018
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
      Left            =   4080
      TabIndex        =   2
      Top             =   1680
      Width           =   3735
   End
End
Attribute VB_Name = "frmVincent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. Johns Housing Project
' Vincent Form
' Written By Kyle Johnson
' 3/22/06
' this form displays additional information about Vincent house including
' a brief description of the house,  and also a floor plan of the house

    
Private Sub cmdBack_Click()
    'navigates from the Vincent page to the options page and also hides the label
    frmOptions.Visible = True
    frmVincent.Visible = False
    lblVincent.Visible = False
    
    
End Sub
    
Private Sub cmdInfoVincent_Click()
    'displays the additional information label
    lblVincent.Visible = True
    
End Sub
    
Private Sub Form_Load()
    'makes the additional info label hidden initially
    lblVincent.Visible = False
End Sub
