VERSION 5.00
Begin VB.Form frmVailresorts 
   BackColor       =   &H00000000&
   Caption         =   "Vail Resorts"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbBack 
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
      Left            =   240
      TabIndex        =   6
      Top             =   9480
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
      Left            =   5520
      ScaleHeight     =   7155
      ScaleWidth      =   7515
      TabIndex        =   5
      Top             =   840
      Width           =   7575
   End
   Begin VB.CommandButton cmd5 
      Height          =   1695
      Left            =   2880
      Picture         =   "frmVailresorts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmd4 
      Height          =   1695
      Left            =   240
      Picture         =   "frmVailresorts.frx":1ACF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmd3 
      Height          =   1695
      Left            =   240
      Picture         =   "frmVailresorts.frx":7FBF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmd2 
      Height          =   1695
      Left            =   240
      Picture         =   "frmVailresorts.frx":944F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmd1 
      Height          =   1695
      Left            =   240
      Picture         =   "frmVailresorts.frx":A25D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   10680
      Width           =   2775
   End
End
Attribute VB_Name = "frmVailresorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmBCresorts(frmBCresorts.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  this form allows the user to get a greater sense of what each
'resort is about. it explains each resort in greater detail than just the price.

Private Sub cmbBack_Click()
    frmVailresorts.Hide 'hides this form
    frmVaillodge.Show 'brings you back to the vaillodge form
End Sub

Private Sub cmd1_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Holiday Inn - Apex Vail"
    picResults.Print 'prints a blank line
    picResults.Print "Conveniently located just minutes from the heart of Vail Village"
    picResults.Print "and less than one mile from the ski lift of Vail Mountain. Free"
    picResults.Print "shuttle operated by hotel during ski season. Free Town of Vail"
    picResults.Print "bus operates year round."

End Sub

Private Sub cmd2_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Lionshead Inn"
    picResults.Print 'prints a blank line
    picResults.Print "The Lionshead Inn offers 52 newly renovated rooms. Enjoy the"
    picResults.Print "refreshing decor of the Inn's lobby and new full-service restaurant."
    picResults.Print "Managed by The Lionshead Inn. Less than 300 yds to slopes."
End Sub

Private Sub cmd3_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Lodge at Avon Center"
    picResults.Print 'prints a blank line
    picResults.Print "Nestled at the base of Beaver Creek Resort, this is a spectacular"
    picResults.Print "enclave location to enjoy all that Colorado has to offer in all"
    picResults.Print "seasons.Managed by Vail Management 10 to 15 minute shuttle ride"
    picResults.Print "to slopes."
End Sub

Private Sub cmd4_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print " Antlers at Vail "
    picResults.Print 'prints a blank line
    picResults.Print "The Antlers at Vail is renowned for its friendly atmosphere and is"
    picResults.Print "located just 150 yards from Vail's gondola and ski school. A $20"
    picResults.Print "million expansion was finished just a few years ago and includes 22"
    picResults.Print "new condominiums, brand-new lobby with courtyard entrance, new"
    picResults.Print "conference rooms, exercise facility, new business center and heated"
    picResults.Print "parking. In early 2004, the Antlers was awarded the Vail Valley"
    picResults.Print "Business of the Year. Starting Spring 2005, Lionshead is undergoing"
    picResults.Print "a transformation that will make it a vibrant, exciting, fun place."
    picResults.Print "It’s business as usual for all mountain operations and the retail,"
    picResults.Print "dining, and lodging establishments during construction."
    picResults.Print "The construction is a reason to visit Lionshead, with Kidstruction"
    picResults.Print "Zone and a sneak peak of the future. Come see for yourself!"
    picResults.Print "Managed by Antlers at Vail. Less than 200 yds to slopes."
End Sub

Private Sub cmd5_Click()
    picResults.Cls 'clears out any info in the pic box
    picResults.Print "Landmark Tower"
    picResults.Print 'prints a blank line
    picResults.Print "Moderately priced condominiums in a choice Lionshead location."
    picResults.Print "Individual furnishings reflect the personalized decor of a vacation"
    picResults.Print "home. Managed by Destination Resorts. Less than 200 yds to slopes."
End Sub

