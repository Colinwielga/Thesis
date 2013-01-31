VERSION 5.00
Begin VB.Form stuffc
   Caption         =   "stuffc"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "stuffc"
   Picture         =   "stuffc.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5
      Caption         =   "Show Name"
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox PicName2
      Height          =   615
      Left            =   480
      ScaleHeight     =   555
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command4
      Caption         =   "Uber Unit?"
      Height          =   855
      Left            =   7560
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3
      Caption         =   "What are my weaknesses?"
      Height          =   855
      Left            =   5520
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2
      Caption         =   "What are my Strengths?"
      Height          =   855
      Left            =   3480
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.PictureBox visuals
      Height          =   3375
      Left            =   3480
      ScaleHeight     =   3315
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   1560
      Width           =   7095
   End
   Begin VB.CommandButton Command1
      Caption         =   "home"
      Height          =   1095
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "stuffc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub thingone_Click()
stuffa.Show
stuffb.Hide
stuffc.Hide

End Sub

Private Sub thingtwo_Click()
visuals.Cls
visuals.Picture = Nothing
visuals.Print "Protoss soldiers and technology have an excellent cost-effectiveness ratio;"
visuals.Print "however, they are quite costly and take some time to train. This leads to a"
visuals.Print "slow start for Protoss players, and a brief period of vulnerability. However,"
visuals.Print "once the economy is rolling, the Protoss can pump out some solid units."
Command3.Visible = True



End Sub

Private Sub photoone_Click()

End Sub

Private Sub thingthree_Click()
visuals.Cls
visuals.Picture = Nothing
visuals.Print "The primary Protoss weakness is inflexibility. The disadvantage of"
visuals.Print "having expensive units isn't the cost, but the commitment that each"
visuals.Print " unit represents. Protoss are the weakest of the three species when it"
visuals.Print "comes to the 'Oh, shoot!' factor. They have a tough time reacting. "
Command4.Visible = True
End Sub

Private Sub thingfour_Click()
 Dim carrier As String

    carrier = "Carrier is just awesome; it has with a metric ton of damage and 'a freakin lot' of health."
   visuals.Picture = LoadPicture(App.Path & "\carrier.JPG")
   visuals.Cls
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print
   visuals.Print



    visuals.Print carrier
End Sub

Private Sub thingfive_Click()
PicName2.Cls
PicName2.Print myname
End Sub

