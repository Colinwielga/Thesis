VERSION 5.00
Begin VB.Form frmAPABook 
   BackColor       =   &H0080FF80&
   Caption         =   "APA"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5720.042
   ScaleMode       =   0  'User
   ScaleWidth      =   9137.03
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdView 
      Caption         =   "View Works Cited"
      Height          =   615
      Left            =   3480
      TabIndex        =   10
      Top             =   5640
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   8715
      TabIndex        =   9
      Top             =   3600
      Width           =   8775
   End
   Begin VB.CommandButton cmdAddCitation 
      Caption         =   "Add Citation"
      Height          =   615
      Left            =   3600
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtCityPublished 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtPublisher 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtAuthorsLastName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Text            =   "Last Name"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtAuthorsMiddleName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Text            =   "M.I."
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtAuthorsFirstName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Text            =   "F.I."
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Go back to main menu"
      Height          =   615
      Left            =   7680
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblStep3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   $"frmAPA.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   4800
      Width           =   8775
   End
   Begin VB.Label lblMostRecent 
      BackColor       =   &H0080FF80&
      Caption         =   "Most recently created citation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label lblStep2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Step 2: Click ""Add Citation"" to add your source to your Works Cited list."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   8775
   End
   Begin VB.Label lblYear 
      BackColor       =   &H0080FF80&
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblCityPublished 
      BackColor       =   &H0080FF80&
      Caption         =   "City Published:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblPublisher 
      BackColor       =   &H0080FF80&
      Caption         =   "Publisher:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H0080FF80&
      Caption         =   "Author's Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H0080FF80&
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblStep1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Step 1: Enter in the following information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmAPABook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form allows the user to input their sources in APA format

Private Sub cmdAddCitation_Click()
    'Adds the information in the text boxes to the parallel arrays
    ctr = ctr + 1
    
    AuthorsLastName(ctr) = txtAuthorsLastName.Text
    AuthorsFirstName(ctr) = txtAuthorsFirstName.Text
    AuthorsMiddleName(ctr) = txtAuthorsMiddleName.Text
    Title(ctr) = txtTitle.Text
    Publisher(ctr) = txtPublisher.Text
    CityPublished(ctr) = txtCityPublished.Text
    Year(ctr) = txtYear.Text
    
    'prints the most recently added citation in the picture box
    picResults.Cls
    picResults.Print AuthorsLastName(ctr) & ", " & AuthorsFirstName(ctr) & ". " & AuthorsMiddleName(ctr) & ". (" & Year(ctr) & "). " & Title(ctr) & ". " & CityPublished(ctr) & ": " & Publisher(ctr) & "."
End Sub


Private Sub cmdMainMenu_Click()
    'moves the user back to the welcome form
    frmAPABook.Hide
    frmWelcome.Show
End Sub

Private Sub cmdView_Click()
    'moves the user to the works cited form to view their bibliography
    frmWorksCited.Show
    frmAPABook.Hide
End Sub
