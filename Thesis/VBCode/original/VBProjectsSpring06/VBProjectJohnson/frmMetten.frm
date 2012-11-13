VERSION 5.00
Begin VB.Form frmMetten 
   BackColor       =   &H00FF0000&
   Caption         =   "Metten Court- Additional Information"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   7845
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
      Height          =   615
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdFloorMetten 
      BackColor       =   &H000000FF&
      Caption         =   "View Floor Plan For Metten by Clicking On Icon"
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
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdInfoMetten 
      BackColor       =   &H000000FF&
      Caption         =   "View General Information for Metten"
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
      Top             =   360
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
      Left            =   5760
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H00FF0000&
      Class           =   "AcroExch.Document.7"
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   600
      OleObjectBlob   =   "frmMetten.frx":0000
      SourceDoc       =   "\\ad\homedir$\Students\KMJOHNSON\My Documents\Metten1st.pdf"
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblMetten 
      BackColor       =   &H000000FF&
      Caption         =   $"frmMetten.frx":3B418
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
End
Attribute VB_Name = "frmMetten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. Johns Housing Project
' metten Form
' Written By Kyle Johnson
' 3/22/06
' this form displays additional information about metten house including
' a brief description of the house,  and also a floor plan of the house



Private Sub cmdBack_Click()
    'takes user from the metten form to the options form
    frmOptions.Visible = True
    frmMetten.Visible = False
    
    lblMetten.Visible = False
    
End Sub
    
    
Private Sub cmdInfoMetten_Click()
    'displays the additional information
    
    lblMetten.Visible = True
    
End Sub
    
Private Sub Form_Load()
    'makes the additional information initially hidden when form is load
    lblMetten.Visible = False
    
End Sub
    
