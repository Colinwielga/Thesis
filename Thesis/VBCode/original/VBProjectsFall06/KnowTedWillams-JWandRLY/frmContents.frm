VERSION 5.00
Begin VB.Form frmContents 
   Caption         =   "Contents"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContentsBack 
      Caption         =   "Click to go back to Introduction Page"
      Height          =   1695
      Left            =   3480
      TabIndex        =   6
      Top             =   6120
      Width           =   3375
   End
   Begin VB.CommandButton cmdStore 
      Caption         =   "Ted's Store"
      Height          =   1695
      Left            =   5040
      TabIndex        =   5
      Top             =   4200
      Width           =   4815
   End
   Begin VB.CommandButton cmdCareerHighs 
      Caption         =   "Ted's Career Highs"
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   4575
   End
   Begin VB.CommandButton cmdBattingTips 
      Caption         =   "Ted's Batting Tips"
      Height          =   1695
      Left            =   5040
      TabIndex        =   3
      Top             =   2280
      Width           =   4815
   End
   Begin VB.CommandButton cmdPictures 
      Caption         =   "Pictures of Ted"
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4575
   End
   Begin VB.CommandButton cmdStats 
      Caption         =   "Ted's Statistics"
      Height          =   1695
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton cmdBiography 
      Caption         =   "Ted's Biography"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "frmContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBattingTips_Click()
frmTedBatTips.Show
frmContents.Hide
End Sub

Private Sub cmdBiography_Click()
frmTedBio.Show
frmContents.Hide
End Sub

Private Sub cmdCareerHighs_Click()
frmTedCareerHighs.Show
frmContents.Hide

End Sub

Private Sub cmdContentsBack_Click()
frmIntro.Show
frmContents.Hide
End Sub

Private Sub cmdPictures_Click()
frmTedPics.Show
frmContents.Hide
End Sub

Private Sub cmdStats_Click()
frmTedStats.Show
frmContents.Hide
End Sub

Private Sub cmdStore_Click()
frmTedStore.Show
frmContents.Hide
End Sub
