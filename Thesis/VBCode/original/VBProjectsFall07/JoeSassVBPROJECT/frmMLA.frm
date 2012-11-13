VERSION 5.00
Begin VB.Form frmMLA 
   Caption         =   "Joe's Citation Creator: MLA"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   2715
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdViewBib 
      Caption         =   "View your Bibliography"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Go back to start"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdWebsite 
      Caption         =   "Webpage"
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdVideo 
      Caption         =   "Videotape or DVD"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncyclo 
      Caption         =   "Encyclopedia"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdPeriodical 
      Caption         =   " Magazine or Newspaper"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdBook 
      Caption         =   "Book"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Line lnDivider 
      X1              =   0
      X2              =   2760
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label lblStep1 
      Alignment       =   2  'Center
      Caption         =   "Step 1: What type of source did you use? Choose from the list below:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmMLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdBook_Click()
frmBook.Show
frmMLA.Hide
End Sub

Private Sub CmdQuit_Click()
frmMLA.Hide
frmWelcome.Show
End Sub



Private Sub cmdViewBib_Click()
frmMLA.Hide
frmWorksCited.Show
End Sub
