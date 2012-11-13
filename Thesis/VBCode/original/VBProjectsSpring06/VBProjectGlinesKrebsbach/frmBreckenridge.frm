VERSION 5.00
Begin VB.Form frmBreckenridge 
   Caption         =   "Breckenridge"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form2"
   Picture         =   "frmBreckenridge.frx":0000
   ScaleHeight     =   10245
   ScaleWidth      =   14025
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
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1575
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
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1575
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
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1575
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
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1575
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9600
      Width           =   1215
   End
   Begin VB.PictureBox piclogo 
      Height          =   1095
      Left            =   11040
      Picture         =   "frmBreckenridge.frx":437D9
      ScaleHeight     =   1035
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   10680
      Width           =   2775
   End
End
Attribute VB_Name = "frmBreckenridge"
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
    frmBreckenridge.Hide
    frmAirline.Show

End Sub


Private Sub cmdFacts_Click()
    MsgBox "Lift Capacity: 37,880 people per hour.  Operating Since: December 16, 1961.  Total Ski/Ride Terrain: 2208 acres / 894 hectares.  Groomed Daily: 600 acres / 241 hectares (29 percent of total terrain)  Bowls: 772 acres / 312 hectares.  Terrain Parks: 25 acres / 10 hectares.  Snowmaking: 565 acres / 228 hectares.  Number of Trails: 146.  Longest Trail: Four O 'Clock - 3.5 miles / 5.6 kilometers", , "Breckenridge Facts"
End Sub

Private Sub cmdLodge_Click()
    frmBreckenridge.Hide
    frmBreckLodge.Show

End Sub

Private Sub cmdtickets_Click()
    frmBreckenridge.Hide
    frmBrecktix.Show

End Sub

Private Sub cmdback_Click()
    frmBreckenridge.Visible = False
    frmContents.Visible = True
End Sub

Private Sub Command4_Click()

End Sub

