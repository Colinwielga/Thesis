VERSION 5.00
Begin VB.Form frmVail 
   Caption         =   "Vail"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form3"
   Picture         =   "frmVail.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdtickets 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lift Tickets"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quick Facts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdLodge 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lodging"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Airfare"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Resorts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   1215
   End
   Begin VB.PictureBox picvail 
      Height          =   1215
      Left            =   120
      Picture         =   "frmVail.frx":26179
      ScaleHeight     =   1155
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   10680
      Width           =   2775
   End
End
Attribute VB_Name = "frmVail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the suser to navigate all the features of the
'beaver creek resort. it allows the user to access forms that search for ticket prices,
'resorts, and airfair. this form also allows the user to read up on quick facts about the
'ski resort

Private Sub cmdAir_Click()
    frmVail.Hide
    frmAirline.Show

End Sub

Private Sub cmdback_Click()
    frmVail.Visible = False
    frmContents.Visible = True
End Sub

Private Sub cmdFacts_Click()
    MsgBox "There are many reasons people choose Vail. First off, the sheer size. It’s the largest single ski resort in North America at 5,289 acres. Second, you can fi nd just about anything here both on and off the mountain—most of all, a good time. A world-renowned Ski & Snowboard School and an array of cool activities and events round out the fun. Simply put, nothing compares to Vail.", , "Vail Facts"
End Sub

Private Sub cmdLodge_Click()
    frmVail.Hide
    frmVaillodge.Show
End Sub

Private Sub cmdtickets_Click()
    frmVail.Hide
    frmvailtix.Show

End Sub



