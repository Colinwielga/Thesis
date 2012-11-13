VERSION 5.00
Begin VB.Form frmBirthstone 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Birthstone"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Go to Main Menu"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3855
      Left            =   3000
      ScaleHeight     =   3795
      ScaleWidth      =   3195
      TabIndex        =   12
      Top             =   480
      Width           =   3255
   End
   Begin VB.CheckBox chkFebruary 
      BackColor       =   &H00C0FFC0&
      Caption         =   "February"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox chkMarch 
      BackColor       =   &H00C0FFC0&
      Caption         =   "March"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox chkApril 
      BackColor       =   &H00C0FFC0&
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox chkMay 
      BackColor       =   &H00C0FFC0&
      Caption         =   "May"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CheckBox chkJune 
      BackColor       =   &H00C0FFC0&
      Caption         =   "June"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox chkJuly 
      BackColor       =   &H00C0FFC0&
      Caption         =   "July"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.CheckBox chkAugust 
      BackColor       =   &H00C0FFC0&
      Caption         =   "August"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox chkSeptember 
      BackColor       =   &H00C0FFC0&
      Caption         =   "September"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CheckBox chkOctober 
      BackColor       =   &H00C0FFC0&
      Caption         =   "October"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox chkNovember 
      BackColor       =   &H00C0FFC0&
      Caption         =   "November"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CheckBox chkDecember 
      BackColor       =   &H00C0FFC0&
      Caption         =   "December"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CheckBox chkJanuary 
      BackColor       =   &H00C0FFC0&
      Caption         =   "January"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmBirthstone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'spell check, checks for errors, Cited: Joe Veeder, Imad Rahal's Problem Description Handout, Lecture 11; this form offers the user a checklist of months, they are to select their birthmonth. It then displays their corresponding birthstone and an image of their birthstone. After experimenting with the different VB Toolbar buttons, we discovered how to use the Checkbox feature.


Private Sub chkApril_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\diamonds.jpg")
        picResults.Print "Diamond"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkAugust_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\peridot.jpg")
        picResults.Print "Peridot"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkDecember_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
       picResults.Picture = LoadPicture(App.Path & "\tanzanite.jpg")
        picResults.Print "Tanzanite"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkFebruary_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\amethyst.jpg")
        picResults.Print "Amethyst"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkJanuary_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\garnet.jpg")
        picResults.Print "Garnet"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub


Private Sub chkJuly_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\ruby.jpg")
        picResults.Print "Ruby"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkJune_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\pearl.jpg")
        picResults.Print "Pearl or Moonstone"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkMarch_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\aquamarine.jpg")
        picResults.Print "Bloodstone or Aquamarine"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkMay_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\emerald.jpg")
        picResults.Print "Emerald"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkNovember_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\citrine.jpg")
        picResults.Print "Citrine"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkOctober_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\opal.jpg")
        picResults.Print "Opal"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub

Private Sub chkSeptember_Click()
    picResults.Cls
    'clears the picture box
    'Cited: Lecture 11
    If True Then
        picResults.Picture = LoadPicture(App.Path & "\sapphire.jpg")
        picResults.Print "Sapphire"
    End If
    'if the particular checkbox is selected, the birthstone and its picture are displayed in the picture box.
    'Cited: http://www.jewelry24seven.com/birthstone_list.htm, http://www.kingsjewelry.com/images/content/gemstones/webgemstones.jpg
End Sub


'This button takes the user to the form frmHome
'Cited: Lecture 18
Private Sub cmdBack_Click()
    frmBirthstone.Hide
    'makes the form frmBirthstone invisible to the user
    frmHome.Show
    'makes the form frmHome visible to the user
End Sub

