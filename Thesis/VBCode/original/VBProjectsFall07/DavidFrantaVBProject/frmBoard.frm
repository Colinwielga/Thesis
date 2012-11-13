VERSION 5.00
Begin VB.Form frmBoard 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   600
   ClientTop       =   1635
   ClientWidth     =   15240
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay Owner"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6840
      TabIndex        =   119
      Top             =   4200
      Width           =   1575
   End
   Begin VB.PictureBox picMoney 
      Height          =   855
      Index           =   3
      Left            =   8520
      ScaleHeight     =   795
      ScaleWidth      =   1755
      TabIndex        =   118
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox txtTransaction 
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   8520
      TabIndex        =   117
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlayerReceive 
      Caption         =   "Receive"
      Enabled         =   0   'False
      Height          =   615
      Index           =   3
      Left            =   9480
      TabIndex        =   115
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton cmdPlayerPay 
      Caption         =   "Pay"
      Enabled         =   0   'False
      Height          =   615
      Index           =   3
      Left            =   8520
      TabIndex        =   114
      Top             =   7920
      Width           =   855
   End
   Begin VB.TextBox txtTransaction 
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   113
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlayerReceive 
      Caption         =   "Receive"
      Enabled         =   0   'False
      Height          =   615
      Index           =   2
      Left            =   5760
      TabIndex        =   111
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton cmdPlayerPay 
      Caption         =   "Pay"
      Enabled         =   0   'False
      Height          =   615
      Index           =   2
      Left            =   4800
      TabIndex        =   110
      Top             =   7920
      Width           =   855
   End
   Begin VB.PictureBox picMoney 
      Height          =   855
      Index           =   2
      Left            =   4800
      ScaleHeight     =   795
      ScaleWidth      =   1755
      TabIndex        =   109
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox txtTransaction 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   8520
      TabIndex        =   108
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtTransaction 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   105
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlayerReceive 
      Caption         =   "Receive"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   9480
      TabIndex        =   104
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdPlayerPay 
      Caption         =   "Pay"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   8520
      TabIndex        =   103
      Top             =   2400
      Width           =   855
   End
   Begin VB.PictureBox picMoney 
      Height          =   855
      Index           =   1
      Left            =   8520
      ScaleHeight     =   795
      ScaleWidth      =   1755
      TabIndex        =   102
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlayerReceive 
      Caption         =   "Receive"
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   5760
      TabIndex        =   101
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdPlayerPay 
      Caption         =   "Pay"
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   4800
      TabIndex        =   100
      Top             =   2400
      Width           =   855
   End
   Begin VB.PictureBox picMoney 
      Height          =   855
      Index           =   0
      Left            =   4800
      ScaleHeight     =   795
      ScaleWidth      =   1755
      TabIndex        =   99
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdChest 
      Caption         =   "Community Chest"
      Enabled         =   0   'False
      Height          =   855
      Left            =   7680
      TabIndex        =   98
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdChance 
      Caption         =   "Chance"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5880
      TabIndex        =   97
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "End Turn"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8520
      TabIndex        =   96
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy Property"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5160
      TabIndex        =   95
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move"
      Enabled         =   0   'False
      Height          =   615
      Index           =   3
      Left            =   12120
      TabIndex        =   94
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll The Dice!"
      Enabled         =   0   'False
      Height          =   615
      Index           =   3
      Left            =   10440
      TabIndex        =   93
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move"
      Enabled         =   0   'False
      Height          =   615
      Index           =   2
      Left            =   3480
      TabIndex        =   92
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll The Dice!"
      Enabled         =   0   'False
      Height          =   615
      Index           =   2
      Left            =   1920
      TabIndex        =   91
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   12120
      TabIndex        =   90
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll The Dice!"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   10440
      TabIndex        =   89
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Property"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6960
      TabIndex        =   88
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move"
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   3480
      TabIndex        =   87
      Top             =   4440
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   735
      Left            =   3480
      ScaleHeight     =   675
      ScaleWidth      =   8115
      TabIndex        =   86
      Top             =   5160
      Width           =   8175
   End
   Begin VB.PictureBox picPlayer 
      Height          =   2895
      Index           =   3
      Left            =   10440
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   85
      Top             =   6720
      Width           =   2775
   End
   Begin VB.PictureBox picPlayer 
      Height          =   2895
      Index           =   2
      Left            =   1920
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   84
      Top             =   6720
      Width           =   2775
   End
   Begin VB.PictureBox picPlayer 
      Height          =   2895
      Index           =   1
      Left            =   10440
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   83
      Top             =   1440
      Width           =   2775
   End
   Begin VB.PictureBox picPlayer 
      Height          =   2895
      Index           =   0
      Left            =   1920
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   82
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Game!"
      Height          =   735
      Left            =   6840
      TabIndex        =   81
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll The Dice!"
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   1920
      TabIndex        =   30
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label43 
      Caption         =   "Amount of Transaction:"
      Height          =   255
      Left            =   8520
      TabIndex        =   116
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label41 
      Caption         =   "Amount of Transaction:"
      Height          =   255
      Left            =   4800
      TabIndex        =   112
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label40 
      Caption         =   "Amount of Transaction:"
      Height          =   255
      Left            =   8520
      TabIndex        =   107
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "Amount of Transaction:"
      Height          =   255
      Left            =   4800
      TabIndex        =   106
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   39
      Left            =   120
      TabIndex        =   80
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   38
      Left            =   120
      TabIndex        =   79
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   37
      Left            =   120
      TabIndex        =   78
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   36
      Left            =   120
      TabIndex        =   77
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   35
      Left            =   120
      TabIndex        =   76
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   34
      Left            =   120
      TabIndex        =   75
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   33
      Left            =   120
      TabIndex        =   74
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   32
      Left            =   120
      TabIndex        =   73
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   31
      Left            =   120
      TabIndex        =   72
      Top             =   9360
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   30
      Left            =   120
      TabIndex        =   71
      Top             =   10680
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   29
      Left            =   1680
      TabIndex        =   70
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   28
      Left            =   3000
      TabIndex        =   69
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   27
      Left            =   4320
      TabIndex        =   68
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   26
      Left            =   5640
      TabIndex        =   67
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   25
      Left            =   6960
      TabIndex        =   66
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   24
      Left            =   8280
      TabIndex        =   65
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   23
      Left            =   9600
      TabIndex        =   64
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   22
      Left            =   10920
      TabIndex        =   63
      Top             =   10560
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   21
      Left            =   12240
      TabIndex        =   62
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   20
      Left            =   13440
      TabIndex        =   61
      Top             =   10560
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   19
      Left            =   13440
      TabIndex        =   60
      Top             =   9480
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   18
      Left            =   13440
      TabIndex        =   59
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   17
      Left            =   13440
      TabIndex        =   58
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   16
      Left            =   13440
      TabIndex        =   57
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   15
      Left            =   13440
      TabIndex        =   56
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   14
      Left            =   13440
      TabIndex        =   55
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   13
      Left            =   13440
      TabIndex        =   54
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   12
      Left            =   13440
      TabIndex        =   53
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   11
      Left            =   13440
      TabIndex        =   52
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label39 
      BackColor       =   &H000080FF&
      Caption         =   "Vincent Court"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12240
      TabIndex        =   51
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label Label38 
      Caption         =   "Chance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10920
      TabIndex        =   50
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label37 
      BackColor       =   &H000080FF&
      Caption         =   "Placid House"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      TabIndex        =   49
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label36 
      BackColor       =   &H000080FF&
      Caption         =   "Maur House"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      TabIndex        =   48
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label35 
      BackColor       =   &H00800080&
      Caption         =   "Sexton Bookstore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6960
      TabIndex        =   47
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label34 
      BackColor       =   &H0000FFFF&
      Caption         =   "Simons Hall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   46
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label33 
      BackColor       =   &H0000FFFF&
      Caption         =   "Peter Engel Science Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4320
      TabIndex        =   45
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label32 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ambulence service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      TabIndex        =   44
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label31 
      BackColor       =   &H0000FFFF&
      Caption         =   "New Science Building"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1680
      TabIndex        =   43
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Free Parking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13440
      TabIndex        =   42
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFF00&
      Caption         =   "Virgil Michel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   41
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFF00&
      Caption         =   "Metten Court"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13440
      TabIndex        =   40
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label27 
      Caption         =   "Community Chest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13440
      TabIndex        =   39
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFF00&
      Caption         =   "Seton Apartments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   38
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label25 
      BackColor       =   &H00800080&
      Caption         =   "IT Sevices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   37
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label24 
      BackColor       =   &H008080FF&
      Caption         =   "Boniface Hall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   36
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label22 
      BackColor       =   &H008080FF&
      Caption         =   "Patrick Hall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   35
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fire Service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   34
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H008080FF&
      Caption         =   "Bernard Hall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      TabIndex        =   33
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label56 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Captured by Life Safety"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   31
      Left            =   120
      TabIndex        =   32
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackColor       =   &H0000C000&
      Caption         =   "Warner Palaestra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   31
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000C000&
      Caption         =   "McNeely Spectrum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Community Chest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C000&
      Caption         =   "Clemens Stadium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800080&
      Caption         =   "Link Bus Service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Chance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "The Quad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Parking Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FF0000&
      Caption         =   "Abby Church and Monestary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   10
      Left            =   13440
      TabIndex        =   21
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13440
      TabIndex        =   20
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   9
      Left            =   12120
      TabIndex        =   19
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tommy Long"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12120
      TabIndex        =   18
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   8
      Left            =   10800
      TabIndex        =   17
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tommy Short"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10800
      TabIndex        =   16
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl1 
      Height          =   375
      Index           =   7
      Left            =   9480
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Chance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9480
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   6
      Left            =   8160
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Mary Hall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8160
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00800080&
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00800080&
      Caption         =   "Alcuin Library "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6840
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   5520
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tuition Increase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Refectory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Community Chest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sexton Dining"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblgo 
      BackColor       =   &H000000FF&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaring many variables.
'there are variables for the dice roll, random numbers, the discription of the property and counters
Dim Roll As Integer, rand1 As Integer, rand2 As Integer, I As Integer, randchance As Integer
Dim Description(1 To 50) As String, J As Integer, Chance(1 To 50) As String
'the Buy command enables the text box for the player whos turn it is. It also disables the Buy and End Turn commands and displays a message box.
Private Sub cmdBuy_Click()
    txtTransaction(Turn - 1).Enabled = True
    cmdBuy.Enabled = False
    cmdOk.Enabled = False
    MsgBox "Enter the value of the property"
End Sub
'This sub generates a random number from one to J, J counter for the length of the data for chance and community chest cards.
'It then displays the text for the corresponding array position
Private Sub cmdChance_Click()
    Randomize
        randchance = Int(J * Rnd) + 1
        picResults.Print Chance(randchance)
        txtTransaction(Turn - 1).Enabled = True
        cmdChance.Enabled = False
        cmdPlayerPay(Turn - 1).Enabled = True
        cmdPlayerReceive(Turn - 1).Enabled = True
        cmdOk.Enabled = True
End Sub
'This is exacly the same as the previous chance command
Private Sub cmdChest_Click()
    Randomize
        randchance = Int(J * Rnd) + 1
        picResults.Print Chance(randchance)
        txtTransaction(Turn - 1).Enabled = True
        cmdChest.Enabled = False
        cmdPlayerPay(Turn - 1).Enabled = True
        cmdPlayerReceive(Turn - 1).Enabled = True
        cmdOk.Enabled = True
End Sub
'This command loads the two files into arrays, one file represents the properties, costs, rents, and avalibility, while the other file represents the possibilities for the chance cards.
'It also sets up the picture boxes on the frmBoard with a heading including the players name.  It displays their begining balance as well.
'It then displays the names of the corresponding player position on the frmPay and frmReceive boards.
Private Sub cmdLoad_Click()
Open App.Path & "\Squares.txt" For Input As #1
Do While Not EOF(1)
    I = I + 1
    Input #1, Description(I), Place(I), Stat(I)
Loop
Close #1
Open App.Path & "\Chance.txt" For Input As #2
Do While Not EOF(2)
    J = J + 1
    Input #2, Chance(J)
Loop
Close #2
picPlayer(0).Print "Properties Owned By "; Player(1)
picPlayer(0).Print "******************************************"
picPlayer(1).Print "Properties Owned By "; Player(2)
picPlayer(1).Print "******************************************"
picPlayer(2).Print "Properties Owned By "; Player(3)
picPlayer(2).Print "******************************************"
picPlayer(3).Print "Properties Owned By "; Player(4)
picPlayer(3).Print "******************************************"
PlayerMoney(1) = 1000
PlayerMoney(2) = 1000
PlayerMoney(3) = 1000
PlayerMoney(4) = 1000
picMoney(0).Print "You have "; FormatCurrency(PlayerMoney(1))
picMoney(1).Print "You have "; FormatCurrency(PlayerMoney(2))
picMoney(2).Print "You have "; FormatCurrency(PlayerMoney(3))
picMoney(3).Print "You have "; FormatCurrency(PlayerMoney(4))
frmPay.lblB.Caption = "Bank"
frmPay.lblP1.Caption = Player(1)
frmPay.lblP2.Caption = Player(2)
frmPay.lblP3.Caption = Player(3)
frmPay.lblP4.Caption = Player(4)
frmReceive.lbl5.Caption = "Bank"
frmReceive.lbl1.Caption = Player(1)
frmReceive.lbl2.Caption = Player(2)
frmReceive.lbl3.Caption = Player(3)
frmReceive.lbl4.Caption = Player(4)

cmdLoad.Enabled = False
cmdRoll(0).Enabled = True
End Sub
'This is a complicated sub that first decides weather the position of the player after the roll is <39 or not.
'It does this to determine if it should subtract 40 from the total, causing the name to appear back at the beginning of the game, instead of on a property that does not exist.
'Within that if statement is another if statement that decides whether it is a property or not and another if statement that decides if it is owned or not.
'It makes these decisions based on a variable in the array that I have assigned a number 0 through 5 to.
'0 for un-owned property, 1 for owned property, 3 for chance, 4 for community chest, 5 for other.

Private Sub cmdMove_Click(Index As Integer)
'Clearing last position
lbl1(Position(Turn)).Caption = ""
'Moving position and starting over at "go" if the player gets to the end
If Position(Turn) + Roll <= 39 Then
    Position(Turn) = Roll + Position(Turn)
    lbl1(Position(Turn)).Caption = Player(Turn)
    picResults.Cls
    picResults.Print Description(Position(Turn) + 1)
    'If it is a property it will have 0 or 1
    If ((Stat(Position(Turn) + 1)) = 0 Or (Stat(Position(Turn) + 1)) = 1) Then
        'if it is un-owned it will have 0
        If ((Stat(Position(Turn) + 1)) = 0) Then
            picResults.Print "You can buy "; Place(Position(Turn) + 1)
            cmdBuy.Enabled = True
            cmdOk.Enabled = True
        Else
            picResults.Print "Pay rent to owner of "; Place(Position(Turn) + 1)
            cmdPay.Enabled = True
        End If
    'for Squares that are not property
    Else
        'For chance and community chest positions
        If ((Stat(Position(Turn) + 1)) = 3 Or (Stat(Position(Turn) + 1)) = 4) Then
            If (Stat(Position(Turn) + 1)) = 3 Then
                cmdChance.Enabled = True
            Else
                cmdChest.Enabled = True
            End If
        'for other positions
        Else
            txtTransaction(Turn - 1).Enabled = True
            cmdPlayerPay(Turn - 1).Enabled = True
            cmdPlayerReceive(Turn - 1).Enabled = True
            cmdOk.Enabled = True
        End If
    End If
'Same as above with the exception of the first line (the player got to the end and is now starting over at the begining of the game)
Else
    Position(Turn) = (Roll + Position(Turn)) - 40
    lbl1(Position(Turn)).Caption = Player(Turn)
    picResults.Cls
    picResults.Print Description(Position(Turn) + 1)
    If ((Stat(Position(Turn) + 1)) = 0 Or (Stat(Position(Turn) + 1)) = 1) Then
        If ((Stat(Position(Turn) + 1)) = 0) Then
            picResults.Print "You can buy "; Place(Position(Turn) + 1)
            cmdBuy.Enabled = True
            cmdOk.Enabled = True
        Else
            picResults.Print "Pay rent to owner of "; Place(Position(Turn) + 1)
            cmdPay.Enabled = True
        End If
    Else
        If ((Stat(Position(Turn) + 1)) = 3 Or (Stat(Position(Turn) + 1)) = 4) Then
            If (Stat(Position(Turn) + 1)) = 3 Then
                cmdChance.Enabled = True
            Else
                cmdChest.Enabled = True
            End If
        Else
            txtTransaction(Turn - 1).Enabled = True
            cmdPlayerPay(Turn - 1).Enabled = True
            cmdPlayerReceive(Turn - 1).Enabled = True
            cmdOk.Enabled = True
        End If
    End If
End If
cmdMove(Turn - 1).Enabled = False


End Sub


'this sub disables all of the current player's buttons and enables the next player to take a turn
'It also checks to see if the current player's account is negative.
Private Sub cmdOk_Click()
'If negative it will end game through a different form
If PlayerMoney(Turn) < 0 Then
    frmLooser.Visible = True
    frmLooser.picLooser.Print Player(Turn); " is the biggest looser in Johnyopoly"
    frmBoard.Visible = False
End If
'checking if there is another player or if the first player gets the next turn
If Turn = NOP Then
    cmdRoll(Turn - NOP).Enabled = True
    cmdMove(Turn - 1).Enabled = False
    Turn = 1
Else
    cmdRoll(Turn).Enabled = True
    cmdMove(Turn - 1).Enabled = False
    Turn = Turn + 1
End If
'disabling the current players'buttons
cmdOk.Enabled = False
cmdBuy.Enabled = False
cmdPlayerPay(Turn - 1).Enabled = False
cmdPlayerReceive(Turn - 1).Enabled = False


End Sub

'instructs player to input rent for property and enables the text box while disabling other commands.
Private Sub cmdPay_Click()
    txtTransaction(Turn - 1).Enabled = True
    cmdBuy.Enabled = False
    cmdOk.Enabled = False
    cmdPay.Enabled = False
    MsgBox "Enter the rent for the property"
End Sub

'disables commands on the frmBoard while introducing a frmPay form
Private Sub cmdPlayerPay_Click(Index As Integer)
    cmdPlayerPay(Turn - 1).Enabled = False
    cmdPlayerReceive(Turn - 1).Enabled = False
    frmPay.Visible = True
    
End Sub

'disables commands on the frmBoard while introducing a frmReceive form
Private Sub cmdPlayerReceive_Click(Index As Integer)
    cmdPlayerPay(Turn - 1).Enabled = False
    cmdPlayerReceive(Turn - 1).Enabled = False
    frmReceive.Visible = True
End Sub
'generates random numbers and adds them and stores their value in a Roll variable
Private Sub cmdRoll_Click(Index As Integer)
Randomize
rand1 = Int(6 * Rnd) + 1
rand2 = Int(6 * Rnd) + 1
Roll = rand1 + rand2
MsgBox "You rolled a " & rand1 & " and a " & rand2 & ".  Move " & Roll & " spaces."
If Turn = 0 Then
    Turn = 1
End If
cmdMove(Turn - 1).Enabled = True
cmdRoll(Turn - 1).Enabled = False

End Sub
'enables the setup board
Private Sub cmdStart_Click()
    frmSetup.Visible = True
    frmBoard.Visible = False
    cmdStart.Enabled = False
    
End Sub

'enables the pay command when text is entered into the textbox.
Private Sub txtTransaction_Change(Index As Integer)
    If txtTransaction(Turn - 1) <> "" Then
        cmdPlayerPay(Turn - 1).Enabled = True
    End If
End Sub
