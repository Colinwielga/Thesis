VERSION 5.00
Begin VB.Form frmSteamboat 
   Caption         =   "Steamboat"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   Picture         =   "frmSteamboat.frx":0000
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   5
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9600
      Width           =   1095
   End
   Begin VB.PictureBox piclogo 
      Height          =   975
      Left            =   360
      Picture         =   "frmSteamboat.frx":23E8D
      ScaleHeight     =   915
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   240
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
Attribute VB_Name = "frmSteamboat"
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
    frmSteamboat.Hide
    frmAirline.Show

End Sub

Private Sub cmdback_Click()
    frmSteamboat.Visible = False
    frmContents.Visible = True
End Sub

Private Sub cmdFacts_Click()
    MsgBox "Tucked away in Colorado is a small town unlike any other. A place where western heritage and genuine friendliness are as honored as the values of a time gone by. And a place where not a whole lot has changed in the last hundred years. It’s just that now, mixed in with the Stetson hats and cowboy boots, is the sophistication of a world-class resort. Our western heritage, six peaks of world-class terrain, and family programs rated the best in the west by SKI Magazine, continue to set Steamboat apart from every other ski resort.  Whether you’re stepping foot into a local pub, one of our charming boutiques, or stepping off the gondola, the reception is the same – warm. But the down home genuine friendliness is only half the reason people choose to vacation here. Nestled 7,000 feet up in the Colorado Rockies, Steamboat is one of North America’s largest ski mountains.", , "Steamboat Facts"

End Sub

Private Sub cmdLodge_Click()
    frmSteamboat.Hide
    frmSteamboatlodge.Show

End Sub

Private Sub cmdtickets_Click()
    frmSteamboat.Hide
    frmSteamboattix.Show

End Sub



