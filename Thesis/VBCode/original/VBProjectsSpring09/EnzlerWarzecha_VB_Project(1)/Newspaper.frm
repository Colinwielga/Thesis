VERSION 5.00
Begin VB.Form frmNews 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   360
      Picture         =   "Newspaper.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   9720
      Picture         =   "Newspaper.frx":3DB6
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(Click on article to read the full story...)"
      BeginProperty Font 
         Name            =   "Myriad Web Pro Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Three escaped convicts thought to have stolen guns"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seacoast Business Owner Arrested On Arizona Fraud Charge"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "24 March 2009"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Police: Alleged Shooter Had A History Of Street Violence"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "                                                                            "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   10215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "          The Daily Journal            "
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   10095
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu return 
         Caption         =   "Return to Menu"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Crime Awareness Project
'frm License
'By: Alex Warzecha & Andrew Enzler
'Written: 3/18/2009

Private Sub Label10_Click()
frmNews.Hide
frmStory2.Show 'Open up a new form to display news article
End Sub

Private Sub Label3_Click()
frmNews.Hide
frmStory1.Show 'Open up a new form to display news article
End Sub

Private Sub Label5_Click()
frmNews.Hide
frmStory3.Show 'Open up a new form to display news article
End Sub
'Offers option to quit
Private Sub quit_Click()
End
End Sub

Private Sub return_Click()
frmNews.Hide
frmHome.Show ' returns to main menu
End Sub


