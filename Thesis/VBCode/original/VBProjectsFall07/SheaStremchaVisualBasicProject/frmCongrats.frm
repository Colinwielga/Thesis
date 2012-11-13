VERSION 5.00
Begin VB.Form frmCongrats 
   BackColor       =   &H0000FF00&
   Caption         =   "Congratulations"
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdRestart 
      BackColor       =   &H00C0C000&
      Caption         =   "Start Over!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   3615
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00C0C000&
      Caption         =   "Go Out and Enjoy the Day"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   3615
   End
   Begin VB.CommandButton cmdNice 
      BackColor       =   &H000000FF&
      Caption         =   "!!!Nice Work!!!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   5895
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   6720
      ScaleHeight     =   6435
      ScaleWidth      =   2955
      TabIndex        =   5
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave and Enjoy Your Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   11280
      Width           =   5775
   End
   Begin VB.PictureBox picFunds 
      BackColor       =   &H0000FF00&
      Height          =   615
      Left            =   6480
      ScaleHeight     =   555
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   9600
      Width           =   3495
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   8
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "And You Still Have This Much Money"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   2
      Top             =   9000
      Width           =   6135
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "You Got Dressed and Ready for the Day!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   1
      Top             =   1800
      Width           =   7335
   End
   Begin VB.Label lblCongrats 
      BackColor       =   &H0000FF00&
      Caption         =   "Congratulations"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmCongrats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'this subroutine takes the user to the survey form
frmCongrats.Visible = False 'Gets rid of current form
frmSurvey.Show 'opens the survery slide
End Sub



Private Sub cmdNice_Click()
'This Button will Show the picture corresponding with the choices that the user made
'It will also show the remainding funds in a seperate picture box
'Prints the User's Name Above the Picture
lblName = Ident
picFunds.Print FormatCurrency(Funds3) 'prints the amount of money from the last form
cmdNice.Visible = False 'allows the picture to be seen
picResults.Picture = LoadPicture(App.Path & "\Images\" & (PicName3) & ".jpg")
End Sub


Private Sub cmdRestart_Click()
'This button will essentially reload the program
cmdNice.Visible = True
frmCongrats.Visible = False
frmName.Visible = True
End Sub
