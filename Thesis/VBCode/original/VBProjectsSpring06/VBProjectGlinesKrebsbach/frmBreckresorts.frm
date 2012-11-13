VERSION 5.00
Begin VB.Form frmBreckresorts 
   BackColor       =   &H00000000&
   Caption         =   "Breck's Resorts"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   13590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Height          =   1695
      Left            =   240
      Picture         =   "frmBreckresorts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmd2 
      Height          =   1695
      Left            =   240
      Picture         =   "frmBreckresorts.frx":6D2B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmd3 
      Height          =   1695
      Left            =   240
      Picture         =   "frmBreckresorts.frx":9424
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   5520
      ScaleHeight     =   7155
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   840
      Width           =   7575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Resorts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   10680
      Width           =   2775
   End
End
Attribute VB_Name = "frmBreckresorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmBreckresorts(frmBreckresorts.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  this form allows the user to get a greater sense of what each
'resort is about. it explains each resort in greater detail than just the price.
Private Sub cmd1_Click()
    picResults.Cls 'clears out any info in the picture box
    picResults.Print "Mountain Thunder Lodge"
    picResults.Print 'prints a blank line
    picResults.Print "Built in 2002, the luxurious Mountain Thunder Lodge features 74"
    picResults.Print "studios, one, two, and three-bedroom condominium suites nestled "
    picResults.Print "in a forested location below Peak 8, just a five-minute walk from"
    picResults.Print "historic Main Street. We are currently building new townhomes at"
    picResults.Print "Mountain Thunder Lodge so there will be construction going on."
    picResults.Print "This will be going on until October and may continue through the"
    picResults.Print "winter season. Managed by Breckenridge Lodging and Hospitality"
    picResults.Print "5 minute shuttle ride to slopes."
End Sub

Private Sub cmd2_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "The Village at Breckenridge Resort"
    picResults.Print 'prints a blank line
    picResults.Print "Where Breckenridge Meets The Mountains. The Village at"
    picResults.Print "Breckenridge Resort is a year-round, western-style family resort"
    picResults.Print "featuring hotel rooms, studios, one, two, and three-bedroom condos,"
    picResults.Print "and also chateaux units. The diverse lodging and 30,000 square feet"
    picResults.Print "of meeting space make us an ideal spot for family vacations, ski"
    picResults.Print "groups, business meetings, conventions, weddings, and family"
    picResults.Print "reunions. Managed by Breckenridge Lodging and Hospitality"
    picResults.Print "Less than 50 yds to slopes."
End Sub

Private Sub cmd3_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "The Chateaux"
    picResults.Print 'prints a clear line
    picResults.Print "Spacious Chateaux condominium suites, located in The Village at"
    picResults.Print "Breckenridge Resort, provide guests with ski-in/ski-out property"
    picResults.Print "access just steps from Main Street. These deluxe accommodations"
    picResults.Print "feature fireplaces, full kitchens and living areas. Chateaux suites"
    picResults.Print "offer the best in both quality and convenience. Managed by "
    picResults.Print "Breckenridge Lodging and Hospitality. Less than 100 yds to slopes."
End Sub

Private Sub cmdback_Click()
    frmBreckresorts.Hide 'hides this form
    frmBreckLodge.Show 'brings you back to the BreckLodge form
End Sub

