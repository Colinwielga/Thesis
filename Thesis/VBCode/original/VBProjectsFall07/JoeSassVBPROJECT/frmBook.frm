VERSION 5.00
Begin VB.Form frmMLABook 
   BackColor       =   &H00FFFF80&
   Caption         =   "MLA"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Go back to main menu"
      Height          =   615
      Left            =   7560
      TabIndex        =   18
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   480
      Width           =   3015
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
      Left            =   5280
      TabIndex        =   7
      Text            =   "First"
      Top             =   480
      Width           =   1215
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
      Left            =   6600
      TabIndex        =   6
      Text            =   "Last"
      Top             =   480
      Width           =   1335
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
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
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
      Left            =   5160
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
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
      Left            =   7800
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdAddCitation 
      Caption         =   "Add Citation"
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   8715
      TabIndex        =   1
      Top             =   3600
      Width           =   8775
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Works Cited"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label lblStep1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
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
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFF80&
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
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00FFFF80&
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
      Left            =   3840
      TabIndex        =   15
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblPublisher 
      BackColor       =   &H00FFFF80&
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
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblCityPublished 
      BackColor       =   &H00FFFF80&
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
      Left            =   3720
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblYear 
      BackColor       =   &H00FFFF80&
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
      Left            =   7200
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblStep2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
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
      Left            =   0
      TabIndex        =   11
      Top             =   1800
      Width           =   8775
   End
   Begin VB.Label lblMostRecent 
      BackColor       =   &H00FFFF80&
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
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label lblStep3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   $"frmBook.frx":0000
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
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Width           =   8775
   End
End
Attribute VB_Name = "frmMLABook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'MLA: adds the users input into the arrays

Private Sub cmdAddCitation_Click()

    'takes the users input from the text boxes and adds it to the arrays
    ctr = ctr + 1
    
    AuthorsLastName(ctr) = txtAuthorsLastName.Text
    AuthorsFirstName(ctr) = txtAuthorsFirstName.Text
    Title(ctr) = txtTitle.Text
    Publisher(ctr) = txtPublisher.Text
    CityPublished(ctr) = txtCityPublished.Text
    Year(ctr) = txtYear.Text
    
    'prints out the last added citation
    picResults.Cls
    picResults.Print AuthorsLastName(ctr) & ", " & AuthorsFirstName(ctr) & ". " & Title(ctr) & ". " & CityPublished(ctr) & ": " & Publisher(ctr) & ", " & Year(ctr) & "."
End Sub


Private Sub cmdMainMenu_Click()
    'brings the user back to the welcome screen
    frmMLABook.Hide
    frmWelcome.Show
End Sub

Private Sub cmdView_Click()
    'brings the user to the works cited form to view their citation
    frmWorksCited.Show
    frmMLABook.Hide
End Sub


