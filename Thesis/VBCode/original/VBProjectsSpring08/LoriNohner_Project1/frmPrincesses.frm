VERSION 5.00
Begin VB.Form frmPrincesses 
   Caption         =   "Princesses!"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   Picture         =   "frmPrincesses.frx":0000
   ScaleHeight     =   7560
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMulan 
      BackColor       =   &H0080FF80&
      Caption         =   "Mulan"
      Height          =   855
      Left            =   7560
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdPocahontas 
      BackColor       =   &H000080FF&
      Caption         =   "Pocahontas"
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdJasmine 
      BackColor       =   &H00FFFF00&
      Caption         =   "Jasmine"
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdBelle 
      BackColor       =   &H0080FFFF&
      Caption         =   "Belle"
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdAriel 
      BackColor       =   &H008080FF&
      Caption         =   "Ariel"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdAurora 
      BackColor       =   &H00FF80FF&
      Caption         =   "Aurora"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdCinderella 
      BackColor       =   &H00FFFF80&
      Caption         =   "Cinderella"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdSnowWhite 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Snow White"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   2895
      Left            =   2760
      ScaleHeight     =   2835
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   4680
      Width           =   4575
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      Height          =   255
      Left            =   8160
      MaskColor       =   &H80000007&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF00FF&
      Caption         =   "Return to Disney Castle"
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "frmPrincesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Project
'Princesses
'Lori Nohner
'Written -March 17, 2008
'Objective- allows user to see individual pictures of the Disney princess they selected.

Private Sub cmdAriel_Click()
    picResults.Picture = LoadPicture(App.Path & "\Ariel2.jpg") 'loads a picture of Ariel into picture box
End Sub

Private Sub cmdAurora_Click()
    picResults.Picture = LoadPicture(App.Path & "\Aurora2.jpg") 'loads a picture of Aurora into picture box
End Sub

Private Sub cmdBelle_Click()
    picResults.Picture = LoadPicture(App.Path & "\Belle2.jpg") 'loads a picture of Belle into picture box
End Sub

Private Sub cmdCinderella_Click()
    picResults.Picture = LoadPicture(App.Path & "\Cinderella2.jpg") 'loads a picture of Cinderella into picture box
End Sub

Private Sub cmdExit_Click()
    End 'quits program
End Sub

Private Sub cmdJasmine_Click()
    picResults.Picture = LoadPicture(App.Path & "\Jasmine2.jpg") 'loads a picture of Jasmine into picture box
End Sub

Private Sub cmdMulan_Click()
    picResults.Picture = LoadPicture(App.Path & "\Mulan2.jpg") 'loads a picture of Mulan into picture box
End Sub

Private Sub cmdPocahontas_Click()
    picResults.Picture = LoadPicture(App.Path & "\Pocahontas2.jpg") 'loads a picture of Pocahontas into picture box
End Sub

Private Sub cmdReturn_Click()
    frmPrincesses.Hide 'hides prnicess page
    frmDisneyCastle.Show 'returns to Disney home page
    
End Sub

Private Sub cmdSnowWhite_Click()
    picResults.Picture = LoadPicture(App.Path & "\Snow-White2.jpg") 'loads a picture of Snow White into picture box
End Sub


Private Sub Form_Load()

End Sub
