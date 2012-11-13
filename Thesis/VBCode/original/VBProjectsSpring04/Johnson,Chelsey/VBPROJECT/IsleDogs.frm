VERSION 5.00
Begin VB.Form IsleDogs 
   BackColor       =   &H0000FF00&
   Caption         =   "Isle of Dogs"
   ClientHeight    =   14580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14580
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
      Height          =   1215
      Left            =   16200
      TabIndex        =   7
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   12720
      TabIndex        =   6
      Top             =   8520
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   6720
      Picture         =   "IsleDogs.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   5640
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Left            =   1320
      Picture         =   "IsleDogs.frx":73C6
      ScaleHeight     =   6555
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   12000
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF00FF&
      Caption         =   $"IsleDogs.frx":FF59
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
      Left            =   6120
      TabIndex        =   5
      Top             =   4200
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   $"IsleDogs.frx":FFF9
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      TabIndex        =   4
      Top             =   3000
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   $"IsleDogs.frx":100EC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      TabIndex        =   3
      Top             =   1560
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "The Most Famous Site, found in the Isle of Dogs district, is the Canary Wharf"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "IsleDogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: IsleDogs(IsleDogs.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: To let the user read the history of the Canary Wharf and also to view pictures of it
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'Returns the user back to the Map of London page so they are able to choose a new district to view and learn the history of.
IsleDogs.Hide
MapLondon.Show
End Sub
