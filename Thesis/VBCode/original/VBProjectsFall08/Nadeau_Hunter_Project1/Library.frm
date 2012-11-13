VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form4"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19110
   FillColor       =   &H00C0C0FF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form4"
   ScaleHeight     =   10890
   ScaleWidth      =   19110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   16320
      TabIndex        =   52
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Menu"
      Height          =   975
      Left            =   13080
      TabIndex        =   51
      Top             =   9720
      Width           =   2295
   End
   Begin VB.PictureBox picOut 
      BackColor       =   &H8000000E&
      Height          =   4815
      Left            =   13080
      ScaleHeight     =   4755
      ScaleWidth      =   5475
      TabIndex        =   50
      Top             =   4680
      Width           =   5535
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Left            =   15360
      ScaleHeight     =   315
      ScaleWidth      =   2715
      TabIndex        =   49
      Top             =   3600
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   15360
      ScaleHeight     =   315
      ScaleWidth      =   2715
      TabIndex        =   48
      Top             =   2520
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   15360
      ScaleHeight     =   315
      ScaleWidth      =   2715
      TabIndex        =   44
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdAuthor 
      Caption         =   "Search by Author Last Name"
      Height          =   495
      Left            =   9240
      TabIndex        =   4
      Top             =   8040
      Width           =   2535
   End
   Begin VB.CommandButton cmdTitle 
      Caption         =   "Search by Title"
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdKeyword 
      Caption         =   "Search by Subject Keyword"
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.PictureBox Picture9 
      Height          =   9735
      Left            =   240
      Picture         =   "Library.frx":0000
      ScaleHeight     =   9675
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   1080
      Width           =   7935
   End
   Begin VB.Label lblKeyword 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Keyword"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   47
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblTitleInfo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   46
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblAuthorInfo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   45
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   18960
      X2              =   18960
      Y1              =   360
      Y2              =   10680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "***********************"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   14400
      TabIndex        =   43
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "*************"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8400
      TabIndex        =   42
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   14280
      TabIndex        =   41
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lblSearch 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Index           =   0
      Left            =   8640
      TabIndex        =   40
      Top             =   360
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   12840
      X2              =   12840
      Y1              =   360
      Y2              =   10680
   End
   Begin VB.Label lblJoseph 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Joseph"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   39
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label lblRilke 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rilke"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   38
      Top             =   10320
      Width           =   615
   End
   Begin VB.Label lblHeindel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Heindel"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   37
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label lblMatheson 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Matheson"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   36
      Top             =   10080
      Width           =   975
   End
   Begin VB.Label lblCarter 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Carter"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   35
      Top             =   9840
      Width           =   735
   End
   Begin VB.Label lblIdel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Idel"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   34
      Top             =   10560
      Width           =   615
   End
   Begin VB.Label lblSMITH 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Smith "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   33
      Top             =   10080
      Width           =   975
   End
   Begin VB.Label lblJocelyn 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Jocelyn"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   32
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label lblGreene 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Greene"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   31
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label lblGebhardt 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gebhardt"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   30
      Top             =   9840
      Width           =   975
   End
   Begin VB.Label lblHodgson 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hodgson"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   29
      Top             =   10560
      Width           =   855
   End
   Begin VB.Label lblAgrippa 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Agrippa"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   28
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label lblSchaya 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Schaya"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   27
      Top             =   10080
      Width           =   855
   End
   Begin VB.Label lblBailey 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bailey"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   26
      Top             =   9600
      Width           =   735
   End
   Begin VB.Label lblLastName 
      BackColor       =   &H00C0FFFF&
      Caption         =   "(last name)"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   25
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Book authors/editors"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   24
      Top             =   8760
      Width           =   3015
   End
   Begin VB.Label lblMsgStars 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Message of the Stars"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   23
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label lblMedZodiac 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Meditations on the Signs of the Zodiac"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   22
      Top             =   5280
      Width           =   4215
   End
   Begin VB.Label lblGeoDescartes 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Geometry of René Descartes"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   21
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Label lblSpinOp 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Spinoza Opera"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   20
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label lbl3BookOccultPhil 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Three Books of Occult Philosophy"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Label lblJewMystEthics 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Jewish Mysticism and Jewish Ethics"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   18
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Label lblKabbalah 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Kabbalah: New Perspectives"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   17
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label lblSonnetsOrpheus 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sonnets to Orpheus"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   16
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lblEsoAst 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Esoteric Astrology"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   15
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblUnivKabbalah 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Universal Meaning of the Kabbalah"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Label lblSufiDoctrine 
      BackColor       =   &H00C0FFFF&
      Caption         =   "An Introduction to Sufi Doctrine"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label lblCloudUnknowing 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Cloud of Unknowing"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   12
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label lblZodiacSoul 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Zodiac and the Soul"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label lblAstFate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Astrology of Fate"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   10
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label lblPhilosophers 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Philosophers"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblMysticism 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mysticism"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Book titles"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblAstrology 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Astrology"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblSubject 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Subject keywords"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblLibrary 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Reuben Brown's Personal Library"
      BeginProperty Font 
         Name            =   "Eras Medium ITC"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Reuben Brown
'Form name: Library
'Author: Nik Nadeau and Zach Hunter
'Date: Nov. 4, 2008
'This form allows user to 'explore' Mr. Brown's personal library.

Dim Ctr As Integer, Keyword(1 To 15) As String, Title(1 To 15) As String, Author(1 To 15) As String

Private Sub cmdClear_Click()
'Clears picture box

picOut.Cls
End Sub

Private Sub cmdAuthor_Click()
'Prompts user to enter author last name; displays info on book by that author.

'Declare variables
Dim Found As Boolean, AR As String, K As Integer

'Assign variables
K = 0
Found = False
AR = InputBox("Enter an author's last name from given list", "Author search")

'Match-stop search
Do Until (Found = True Or K > 14)
    K = K + 1
    If AR = Author(K) Then
        Found = True
    End If
Loop

'If there's a match, print corresponding book cover and bibliographic information
If Found Then
    Select Case AR
        Case "Joseph"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Dan.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
    
        Case "Idel"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Idel.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Schaya"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Schaya.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
       
        Case "Hodgson"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Cloud.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Rilke"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Rilke.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Matheson"""
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Burckhardt.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Bailey"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Bailey.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Carter"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Carter.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Greene"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Greene.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
        
        Case "Heindel"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Heindel.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Jocelyn"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Jocelyn.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Smith"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Descartes.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Gebhardt"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Spinoza.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
        Case "Agrippa"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Agrippa.jpg")
            Picture1.Print Title(Ctr)
            Picture2.Print AR
            Picture3.Print Keyword(Ctr)
            
    End Select
    
End If

End Sub

Private Sub cmdKeyword_Click()
'Prompts user to enter keyword; displays info on book categorized by that keyword.

Dim KW As String

'Assign variable
KW = LCase(InputBox("Enter a keyword from given list", "Keyword search"))

'Print corressponding book titles based on keyword input
Select Case KW
    Case "mysticism"
        picOut.Cls
        picOut.Print "Check out these titles:"
        picOut.Print
        picOut.Print "Jewish Mysticism and Jewish Ethics"
        picOut.Print "Kabbalah: New Perspectives"
        picOut.Print "The Universal Meaning of the Kabbalah"
        picOut.Print "The Cloud of Unknowing"
        picOut.Print "Sonnets to Orpheus"
        picOut.Print "An Introduction to Sufi Doctrine"
    Case "astrology"
        picOut.Cls
        picOut.Print "Check out these titles:"
        picOut.Print
        picOut.Print "Esoteric Astrology"
        picOut.Print "The Zodiac and the Soul"
        picOut.Print "The Astrology of Fate"
        picOut.Print "The Message of the Stars"
        picOut.Print "Meditations on the Signs of the Zodiac"
    Case "philosophers"
        picOut.Cls
        picOut.Print "Check out these titles:"
        picOut.Print
        picOut.Print "The Geometry of René Descartes"
        picOut.Print "Spinoza Opera"
        picOut.Print "Three Books of Occult Philosophy"
End Select

End Sub


Private Sub cmdMenu_Click()
'Back to Menu

Form4.Hide
Form1.Show
End Sub

Private Sub cmdQuit_Click()
'Quit

End
End Sub

Private Sub cmdTitle_Click()
'Prompts user to enter title; displays info on book by that title.

'Declare variables
Dim Found As Boolean, TL As String, K As Integer

'Assign variables
K = 0
Found = False
TL = InputBox("Enter a title from given list, using appropriate capital letters", "Title search")

'Match stop search
Do Until (Found = True Or K > 14)
    K = K + 1
    If TL = Title(K) Then
        Found = True
    End If
Loop

'If there's a match, print corresponding book cover and bibliographic information
If Found Then
    Select Case TL
        Case "Jewish Mysticism and Jewish Ethics"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Dan.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
    
        Case "Kabbalah: New Perspectives"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Idel.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "The Universal Meaning of the Kabbalah"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Schaya.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
       
        Case "The Cloud of Unknowing"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Cloud.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "Sonnets to Orpheus"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Rilke.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "An Introduction to Sufi Doctrine"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Burckhardt.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "Esoteric Astrology"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Bailey.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "The Zodiac and the Soul"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Carter.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "The Astrology of Fate"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Greene.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
        
        Case "The Message of the Stars"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Heindel.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "Meditations on the Signs of the Zodiac"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Jocelyn.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "The Geometry of René Descartes"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Descartes.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "Spinoza Opera"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Spinoza.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
        Case "Three Books of Occult Philosophy"
            picOut.Cls
            picOut.Picture = LoadPicture(App.Path & "\Agrippa.jpg")
            Picture1.Print TL
            Picture2.Print Author(Ctr)
            Picture3.Print Keyword(Ctr)
            
    End Select
    
End If

End Sub

Private Sub Form_Load()
'Load array at start of form load

'Open data file BooksSearch.txt
Open App.Path & "\BooksSearch.txt" For Input As #1

'Set counter to zero
Ctr = 0

'Read array from data file BooksSearch.txt
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Keyword(Ctr), Title(Ctr), Author(Ctr)
Loop

'Close data file
Close #1

End Sub

