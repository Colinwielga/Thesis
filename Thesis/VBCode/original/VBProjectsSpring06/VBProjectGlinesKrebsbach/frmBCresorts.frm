VERSION 5.00
Begin VB.Form frmBCresorts 
   BackColor       =   &H00000000&
   Caption         =   "Beaver Creek Resorts"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
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
      Left            =   480
      TabIndex        =   9
      Top             =   9240
      Width           =   1335
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
      Left            =   5640
      ScaleHeight     =   7155
      ScaleWidth      =   7515
      TabIndex        =   8
      Top             =   720
      Width           =   7575
   End
   Begin VB.CommandButton cmd8 
      Height          =   1695
      Left            =   3240
      Picture         =   "frmBCresorts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton cmd7 
      Height          =   1695
      Left            =   3240
      Picture         =   "frmBCresorts.frx":5367
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmd6 
      Height          =   1695
      Left            =   3240
      Picture         =   "frmBCresorts.frx":A14B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmd5 
      Height          =   1695
      Left            =   3240
      Picture         =   "frmBCresorts.frx":B936
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmd4 
      Height          =   1695
      Left            =   360
      Picture         =   "frmBCresorts.frx":1265F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton cmd3 
      Height          =   1695
      Left            =   360
      Picture         =   "frmBCresorts.frx":13696
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmd2 
      Height          =   1695
      Left            =   360
      Picture         =   "frmBCresorts.frx":1966F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmd1 
      Height          =   1695
      Left            =   360
      Picture         =   "frmBCresorts.frx":1AA8E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   10560
      Width           =   2775
   End
End
Attribute VB_Name = "frmBCresorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()

End Sub

'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmBCresorts(frmBCresorts.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  this form allows the user to get a greater sense of what each
'resort is about. it explains each resort in greater detail than just the price.
Private Sub cmd1_Click()
    picResults.Cls 'clears out all info in the picture box
    picResults.Print "Bachlor Gulch Condominiums"
    picResults.Print
    picResults.Print "Premium two and three bedroom condominiums located in Bachelor Gulch."
    picResults.Print "Enjoy ski-in/ski-out convenience, unparalleled service and resort"
    picResults.Print "amenities.  Managed by Vail/Beaver Creek Resort Properties."
    picResults.Print "Ski in/ski out."
End Sub

Private Sub cmd2_Click()
    picResults.Cls 'clears out all info in the picture box
    picResults.Print "Chapel Square"
    picResults.Print
    picResults.Print "Avon's newest condominiums, Chapel Square, are located in the heart"
    picResults.Print "of Avon within walking distance to numerous dining and shopping"
    picResults.Print "options, also within minutes of Vail and Beaver Creek Resorts. "
    picResults.Print "The location has something for everyone.  Managed by "
    picResults.Print "Vail/Beaver Creek Resort Properties.  10 to 15 minute shuttle "
    picResults.Print "ride to slopes."
End Sub

Private Sub cmd3_Click()
    picResults.Cls 'clears all info in pic box
    picResults.Print "Elkhorn Lodge"
    picResults.Print
    picResults.Print "Ski in/ski out from mid December to early April. Elkhorn Lodge "
    picResults.Print "offers a year-round experience with its location next to the "
    picResults.Print "Beaver Creek Golf Clubhouse and a chair lift out the back door. "
    picResults.Print "Within minutes to world class dining and luxury shopping options."
    picResults.Print "Managed by Vail/Beaver Creek Resort Properties. Ski in/ski out."
End Sub

Private Sub cmd4_Click()
    picResults.Cls 'clears all info in the pic box
    picResults.Print "Saddleridge"
    picResults.Print
    picResults.Print "A premier mountain resort offering villas furnished with western"
    picResults.Print "antiques. Luxurious two bedroom villas, inspired with tasteful "
    picResults.Print "accents of Ralph Lauren fabrics and linens, offering full gourmet"
    picResults.Print "kitchens, spacious living areas, and garage parking. Managed by "
    picResults.Print "Vail/Beaver Creek Resort Properties.5 minute shuttle ride to slope"
End Sub

Private Sub cmd5_Click()
    picResults.Cls 'clears all info in the pic box
    picResults.Print "Seasons at Avon"
    picResults.Print
    picResults.Print "The best things about the Seasons at Avon are the walk-out access"
    picResults.Print "to all of Avon's amenities and every unit has a mountain view!"
    picResults.Print "Managed by Vail/Beaver Creek Resort Properties. 10 min ride"
End Sub

Private Sub cmd6_Click()
    picResults.Cls 'clears all info in the pic box
    picResults.Print "Snow Cloud"
    picResults.Print
    picResults.Print "Located at the base of the Bachelor Gulch lift, Snow Cloud is an"
    picResults.Print "ideal ski in/ski out luxury condominium property. The Ritz Carlton"
    picResults.Print "is located adjacent to this property and guests have full access"
    picResults.Print "to the pool. There are several hot tubs on the Snow Cloud property"
    picResults.Print "as well. These oversized and comfortable condos are the perfect"
    picResults.Print "place to spend your vacation. Enjoy village to village skiing"
    picResults.Print "between Bachelor Gulch, Beaver Creek and Arrowhead Villages."
    picResults.Print "Managed by Vail/Beaver Creek Resort Properties.Ski in/ski out."
End Sub

Private Sub cmd7_Click()
    picResults.Cls 'clears all info in the pic box
    picResults.Print "Townsend Place"
    picResults.Print
    picResults.Print "Ski in/ski out from mid December to early April. Townsend Place"
    picResults.Print "is centrally located next to a quietly meandering creek in Beaver"
    picResults.Print "Creek Resort. The condominiums offer a wood burning fireplace,"
    picResults.Print "washer/dryer, balcony, fully equipped kitchen, and a value option"
    picResults.Print "in the heart of Beaver Creek."
    picResults.Print "Managed by Vail/Beaver Creek Resort Properties.Ski in/ski out."
    End Sub

Private Sub cmd8_Click()
    picResults.Cls 'clears all info in the pic box
    picResults.Print "Lodge at Brookside"
    picResults.Print
    picResults.Print "The Lodge at Brookside is the ideal location next to the meandering"
    picResults.Print "river, with views of the ski slopes. Enjoy abundant year-round"
    picResults.Print "recreation while staying in our oversized condominiums complete"
    picResults.Print "with gourmet kitchens and all the amenities of home."; Managed; By; ""
    picResults.Print "Vail/Beaver Creek Resort Properties. 10 to 15 minute shuttle ride"
End Sub
Private Sub cmdback_Click()
frmBCresorts.Hide 'hides this form
frmBCLodge.Show 'shows the BClodge form
End Sub
