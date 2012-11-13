VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
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
   Begin VB.PictureBox picresults 
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form2.Hide
Form3.Hide

End Sub

Private Sub Command2_Click()
picresults.Cls
picresults.Picture = Nothing
picresults.Print "Protoss soldiers and technology have an excellent cost-effectiveness ratio;"
picresults.Print "however, they are quite costly and take some time to train. This leads to a"
picresults.Print "slow start for Protoss players, and a brief period of vulnerability. However,"
picresults.Print "once the economy is rolling, the Protoss can pump out some solid units."
Command3.Visible = True



End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Command3_Click()
picresults.Cls
picresults.Picture = Nothing
picresults.Print "The primary Protoss weakness is inflexibility. The disadvantage of"
picresults.Print "having expensive units isn't the cost, but the commitment that each"
picresults.Print " unit represents. Protoss are the weakest of the three species when it"
picresults.Print "comes to the 'Oh, shoot!' factor. They have a tough time reacting. "
Command4.Visible = True
End Sub

Private Sub Command4_Click()
 Dim carrier As String

    carrier = "Carrier is just awesome; it has with a metric ton of damage and 'a freakin lot' of health."
   picresults.Picture = LoadPicture(App.Path & "\carrier.JPG")
   picresults.Cls
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
   picresults.Print
  
  
  
    picresults.Print carrier
End Sub

Private Sub Command5_Click()
PicName2.Cls
PicName2.Print myname
End Sub

