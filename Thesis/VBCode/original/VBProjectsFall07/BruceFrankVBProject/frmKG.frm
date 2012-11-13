VERSION 5.00
Begin VB.Form frmKG 
   BackColor       =   &H0000C0C0&
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   3795
   ClientTop       =   1890
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   9060
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7440
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdfarewell 
      Caption         =   "Farewell KG"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7440
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdPic 
      Caption         =   "Check Out Some KG Pics"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7440
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.PictureBox picStats 
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdStats 
      Caption         =   "Check out Some of KG's Stats"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdBio 
      Caption         =   "BIO"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblKG 
      BackColor       =   &H0000C0C0&
      Caption         =   "KEVIN GARNETT: The Greatest Timberwolf EVER!"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   1680
      Picture         =   "frmKG.frx":0000
      Top             =   960
      Width           =   5625
   End
End
Attribute VB_Name = "frmKG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is devoted to KG and gives general info such as a bio, stats, pictures, and a farewell due to his recent trade to the Celtics


Private Sub cmdBio_Click()
'This command button displays a bio.  A message box is used to display this info.
MsgBox "Born: 19-May-1976, Height: 6-11, Weight: 220  lbs., College: Farragut Academy HS (IL)Years Pro: 12. Kevin Garnett grew up in Mauldin, South Carolina where he became a high school b-ball star and was selected the state's Mr. Basketball in 1994. Garnett moved to Chicago for his senior year and, in 1995, was drafted by the Minnesota Timberwolves - becoming the first NBA player in 20 years to be picked right out of high school. Kevin Garnett decided to skip college after realizing he would be a high pick in the draft and would be in line for some major moola.  Kevin Garnett signed a multi-million dollar contract with the T-Wolves when he was 19, and was named to the All-Star team after his second season. The rest is history. Kevin Garnett is engaged to his long-time girlfriend, Brandi Padilla."

End Sub

Private Sub cmdfarewell_Click()
'This command button gives a farewell to KG because he was traded to Boston.  It uses the form of a message box
MsgBox "Recently KG was traded to the Boston Celtics.  The Timberwolves are sad to see him go but happy that he has a legitimate chance for an NBA championship.  Farewell KG."
End Sub

Private Sub cmdPic_Click()
'This command button allows the user to view pictures of KG.  It hides the KG page and redirects the user to as KG picture
frmKG.Visible = False
frmKGpic1.Visible = True

End Sub

Private Sub cmdreturn_Click()
'This command button returns the user to the main page form and away from the KG form
frmKG.Visible = False
frmMainPage.Visible = True

End Sub

Private Sub cmdStats_Click()
'This form shows KG's career stats for rebounding and points
'the picture box is cleared so that if a user prsses the command button more than once the same info is not displayed twice

picStats.Cls
picStats.Print "Year   PPG     RPG"
picStats.Print "'96    10.4    6.3"
picStats.Print "'97    17.0    8.1"
picStats.Print "'98    18.5    9.6"
picStats.Print "'99    20.8    10.4"
picStats.Print "'00    22.9    11.8"
picStats.Print "'01    22.0    11.4"
picStats.Print "'02    21.2    12.1"
picStats.Print "'03    23.0    13.5"
picStats.Print "'04    24.2    13.9"
picStats.Print "'05    22.2    13.5"
picStats.Print "'06    21.8    12.7"
picStats.Print "'07    22.4    12.8"
picStats.Print "'08    22.0    20.0"

End Sub
