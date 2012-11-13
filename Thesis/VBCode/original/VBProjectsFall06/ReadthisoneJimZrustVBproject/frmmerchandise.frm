VERSION 5.00
Begin VB.Form frmmerchandise 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturnorder 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Order Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6600
      Width           =   3015
   End
   Begin VB.PictureBox Picture6 
      Height          =   2415
      Left            =   6120
      Picture         =   "frmmerchandise.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   6120
      Picture         =   "frmmerchandise.frx":5388
      ScaleHeight     =   1995
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
      Begin VB.PictureBox Picture5 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   1935
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   3615
      Left            =   0
      Picture         =   "frmmerchandise.frx":6D0B
      ScaleHeight     =   3555
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   4200
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   0
      Picture         =   "frmmerchandise.frx":9CCC
      ScaleHeight     =   1995
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   0
      Picture         =   "frmmerchandise.frx":E4E7
      ScaleHeight     =   1995
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblposter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Poster- $5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label lblhat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hat- $15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblshirt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tee-Shirt- $17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblsweat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sweatshirt- $50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lbljacket 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jacket- $100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "frmmerchandise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Title: Minnesota Vikings Fan Page

'Form Name: Merchandise

'Written by Jim Zrust

'Date: November 1, 2006

'Form Objective: this form was created to allow the user to see what products they could buy
'at the team store and their corresponding price.  it was a simple form that only had
'a button to return the user to the team store

Private Sub cmdreturnorder_Click()
frmmerchandise.Hide
frmorder.Show
End Sub

