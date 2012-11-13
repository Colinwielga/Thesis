VERSION 5.00
Begin VB.Form frmstadiumpictures 
   BackColor       =   &H000000C0&
   Caption         =   "Pictures"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdpic5 
      Caption         =   "Picture 5"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdpic2 
      Caption         =   "Picture 2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdpic3 
      Caption         =   "Picture 3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdpic4 
      Caption         =   "Picture 4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "Picture 1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   10125
      Left            =   0
      Picture         =   "frmstadiumpictures.frx":0000
      Top             =   720
      Width           =   13500
   End
   Begin VB.Image Image5 
      Enabled         =   0   'False
      Height          =   7200
      Left            =   480
      Picture         =   "frmstadiumpictures.frx":1BCF66
      Top             =   960
      Width           =   9600
   End
   Begin VB.Image Image4 
      Enabled         =   0   'False
      Height          =   2835
      Left            =   4200
      Picture         =   "frmstadiumpictures.frx":252FB4
      Top             =   3120
      Width           =   2250
   End
   Begin VB.Image Image3 
      Enabled         =   0   'False
      Height          =   3750
      Left            =   2280
      Picture         =   "frmstadiumpictures.frx":260D7E
      Top             =   2760
      Width           =   5250
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   4380
      Left            =   1680
      Picture         =   "frmstadiumpictures.frx":28B964
      Top             =   2280
      Width           =   7500
   End
   Begin VB.Image Image1 
      Height          =   2715
      Left            =   2640
      Picture         =   "frmstadiumpictures.frx":2D2E52
      Top             =   2760
      Width           =   5250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pictures of Rosenblatt Stadium"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frmstadiumpictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'College World Series.(NCAACollegeWorldSeries.vbp)

'Form name: frmstadiumpictures; Form caption: Pictures

'Author: Sam Dorr

'Date written: March 25, 2007

' Form Objective: The objective of frmstadiumpictures is to display five seperate pictures
'                   to see what Rosenblatt Stadium, the home of the CWS, looks like.  The
'                   pictures are displayed one by one at the users request by making the
'                   visibility of the bitmap true or false.

Option Explicit

Private Sub cmdback_Click()
    Image1.Visible = False 'denies picture visibility
    Image2.Visible = False 'denies picture visibility
    Image3.Visible = False 'denies picture visibility
    Image4.Visible = False 'denies picture visibility
    Image5.Visible = False 'denies picture visibility
    Image6.Visible = True  'allows picture visibility
    frmstadiumpictures.Hide
    frmhome.Show
End Sub

Private Sub cmdpic1_Click()
    Image1.Visible = True
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
End Sub

Private Sub cmdpic2_Click()
    Image1.Visible = False
    Image2.Visible = True
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
End Sub

Private Sub cmdpic3_Click()
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = True
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
End Sub

Private Sub cmdpic4_Click()
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = True
    Image5.Visible = False
    Image6.Visible = False
End Sub

Private Sub cmdpic5_Click()
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = True
    Image6.Visible = False
End Sub
