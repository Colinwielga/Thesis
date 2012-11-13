VERSION 5.00
Begin VB.Form frmCharacter 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdA 
      BackColor       =   &H80000015&
      Caption         =   "Alice"
      Height          =   800
      Left            =   11880
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8160
      Width           =   2500
   End
   Begin VB.CommandButton cmdK 
      BackColor       =   &H80000015&
      Caption         =   "Select"
      Height          =   800
      Left            =   4560
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8040
      Width           =   2500
   End
   Begin VB.CommandButton cmdAD 
      BackColor       =   &H80000015&
      Caption         =   "Select"
      Height          =   800
      Left            =   11880
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   2500
   End
   Begin VB.CommandButton cmdBO 
      BackColor       =   &H80000015&
      Caption         =   "Select"
      Height          =   800
      Left            =   4560
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   2500
   End
   Begin VB.PictureBox picAlice 
      Height          =   2500
      Left            =   8640
      ScaleHeight     =   2445
      ScaleWidth      =   2940
      TabIndex        =   4
      Top             =   6360
      Width           =   3000
   End
   Begin VB.PictureBox picArthur 
      Height          =   2500
      Left            =   8640
      ScaleHeight     =   2445
      ScaleWidth      =   2940
      TabIndex        =   3
      Top             =   3240
      Width           =   3000
   End
   Begin VB.PictureBox picMartha 
      Height          =   2500
      Left            =   1320
      ScaleHeight     =   2445
      ScaleWidth      =   2940
      TabIndex        =   2
      Top             =   6360
      Width           =   3000
   End
   Begin VB.PictureBox picBarack 
      Height          =   2500
      Left            =   1320
      ScaleHeight     =   2445
      ScaleWidth      =   2940
      TabIndex        =   1
      Top             =   3240
      Width           =   3000
   End
   Begin VB.Label lblAlice 
      BackColor       =   &H00800080&
      Caption         =   "Alice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   8
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label lblK 
      BackColor       =   &H00800080&
      Caption         =   "Martha Stewart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4560
      TabIndex        =   7
      Top             =   7200
      Width           =   1170
   End
   Begin VB.Label lblArthur 
      BackColor       =   &H00800080&
      Caption         =   "Arthur Dent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11880
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblBarack 
      BackColor       =   &H00800080&
      Caption         =   "Barack Obama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblCharacter 
      BackColor       =   &H00800080&
      Caption         =   "Choose Your Character!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   11415
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmCharacter
'Author:  Peter Woodruff
'Date Written:  3-20-09
'Purpose:  This form allows the user to select his/her character.

Option Explicit

Private Sub cmdA_Click()
    
    'Select Alice
    frmCharacter.Visible = False
    frmAlice.Visible = True

    
End Sub

Private Sub cmdAD_Click()
    
    'Select Arthur
    frmCharacter.Visible = False
    frmArthur.Visible = True

    
End Sub

Private Sub cmdBO_Click()

    'Select Obama, sets a number for name and life arrays
    frmCharacter.Visible = False
    frmBarack.Visible = True

    
End Sub

Private Sub cmdK_Click()
    
    'Select Martha
    frmCharacter.Visible = False
    frmMartha.Visible = True

    
End Sub

Private Sub Form_Load()

    'Loads pictures of everyone
    
    picBarack.Picture = LoadPicture(App.Path & "\BO.bmp")
    picArthur.Picture = LoadPicture(App.Path & "\AD.gif")
    picMartha.Picture = LoadPicture(App.Path & "\MS.bmp")
    picAlice.Picture = LoadPicture(App.Path & "\AW.jpg")

    
End Sub

Private Sub lblCharacter_Click()

End Sub
