VERSION 5.00
Begin VB.Form frmTransport 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Wilderness Outfitters"
   ClientHeight    =   9705
   ClientLeft      =   3360
   ClientTop       =   945
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   Picture         =   "frmTransport.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   5430
   Begin VB.CommandButton cmdPreviousTransport 
      Caption         =   "Previous Page"
      Height          =   495
      Left            =   2160
      TabIndex        =   47
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubmitTransport 
      Caption         =   "Submit"
      Height          =   495
      Left            =   480
      TabIndex        =   46
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   22
      Left            =   480
      TabIndex        =   21
      Text            =   "0"
      Top             =   8640
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   21
      Left            =   480
      TabIndex        =   20
      Text            =   "0"
      Top             =   8280
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   20
      Left            =   480
      TabIndex        =   19
      Text            =   "0"
      Top             =   7920
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   19
      Left            =   480
      TabIndex        =   18
      Text            =   "0"
      Top             =   7560
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   18
      Left            =   480
      TabIndex        =   17
      Text            =   "0"
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   17
      Left            =   480
      TabIndex        =   16
      Text            =   "0"
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   16
      Left            =   480
      TabIndex        =   15
      Text            =   "0"
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   15
      Left            =   480
      TabIndex        =   14
      Text            =   "0"
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   14
      Left            =   480
      TabIndex        =   13
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   13
      Left            =   480
      TabIndex        =   12
      Text            =   "0"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   12
      Left            =   480
      TabIndex        =   11
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   11
      Left            =   480
      TabIndex        =   10
      Text            =   "0"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   10
      Left            =   480
      TabIndex        =   9
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   9
      Left            =   480
      TabIndex        =   8
      Text            =   "0"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   7
      Text            =   "0"
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   6
      Text            =   "0"
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   5
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   4
      Text            =   "0"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   3
      Text            =   "0"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   2
      Text            =   "0"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   1
      Text            =   "0"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtTransport 
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Text            =   "0"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label47 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   70
      Top             =   8640
      Width           =   1000
   End
   Begin VB.Label Label46 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   69
      Top             =   8280
      Width           =   1000
   End
   Begin VB.Label Label45 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$450.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   68
      Top             =   7920
      Width           =   1000
   End
   Begin VB.Label Label44 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$300.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   67
      Top             =   7560
      Width           =   1000
   End
   Begin VB.Label Label43 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$40.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   66
      Top             =   7200
      Width           =   1000
   End
   Begin VB.Label Label42 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$110.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   65
      Top             =   6840
      Width           =   1000
   End
   Begin VB.Label Label41 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$25.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   64
      Top             =   6480
      Width           =   1000
   End
   Begin VB.Label Label40 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$30.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   63
      Top             =   6120
      Width           =   1000
   End
   Begin VB.Label Label39 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$35.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   62
      Top             =   5760
      Width           =   1000
   End
   Begin VB.Label Label38 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$70.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   61
      Top             =   5400
      Width           =   1000
   End
   Begin VB.Label Label37 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$45.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   60
      Top             =   5040
      Width           =   1000
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$30.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   59
      Top             =   4680
      Width           =   1000
   End
   Begin VB.Label Label35 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$35.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   58
      Top             =   4320
      Width           =   1000
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$125.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   57
      Top             =   3960
      Width           =   1000
   End
   Begin VB.Label Label33 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$80.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   56
      Top             =   3600
      Width           =   1000
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$80.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   55
      Top             =   3240
      Width           =   1000
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$40.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   54
      Top             =   2880
      Width           =   1000
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$40.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   53
      Top             =   2520
      Width           =   1000
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$25.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   52
      Top             =   2160
      Width           =   1000
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$20.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   51
      Top             =   1800
      Width           =   1000
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$25.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   50
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$45.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   49
      Top             =   1080
      Width           =   1000
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   48
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Destinations"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   45
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportation"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   44
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Beaverhouse"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   43
      Top             =   8640
      Width           =   2505
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Atikokan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   42
      Top             =   8280
      Width           =   2505
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Saganaga"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   41
      Top             =   7920
      Width           =   2505
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sawbill Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   40
      Top             =   7560
      Width           =   2505
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Lake Vermilion (Rice Bay)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   39
      Top             =   7200
      Width           =   2505
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Crane Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   38
      Top             =   6840
      Width           =   2505
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Wood Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   37
      Top             =   6480
      Width           =   2505
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "South Kawishiwi Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   36
      Top             =   6120
      Width           =   2505
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Snowbank Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   35
      Top             =   5760
      Width           =   2505
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nina Moose Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   34
      Top             =   5400
      Width           =   2505
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mudro Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   33
      Top             =   5040
      Width           =   2505
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Moose Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   32
      Top             =   4680
      Width           =   2505
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Lake One"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   31
      Top             =   4320
      Width           =   2505
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Kawishiwi Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   30
      Top             =   3960
      Width           =   2505
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Isabella Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   29
      Top             =   3600
      Width           =   2505
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Indian Sioux River"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   3240
      Width           =   2505
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hegman Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   27
      Top             =   2880
      Width           =   2505
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gabbro Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   2520
      Width           =   2505
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Farm Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   25
      Top             =   2160
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fall Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   1800
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Burntside Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   23
      Top             =   1440
      Width           =   2505
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Angleworm Lake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   1080
      Width           =   2505
   End
End
Attribute VB_Name = "frmTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wilderness Outfitters Partial Outfitting
'Justin Bork
'March, 2009
'
'frmTransport
'The purpose of this form is to allow the user to get hauled to varios
'available destinations by vehicle.
Private Sub cmdPreviousTransport_Click()
    frmTransport.Hide
    frmStartup.Show
End Sub

Private Sub cmdSubmitTransport_Click()
    'This button first takes note of what items are requested and writes that
    'into the user's text file.
    'It then prints the user's selections.
    Dim i As Integer
    subtotal1 = 0
    subtotal2 = 0
    subtotal3 = 0
    grandTotal = 0
    i = 0
    
    frmDisplay.picResults.Cls
    
    For pos = 65 To 86
        i = i + 1
        Requests(pos) = txtTransport(i).Text
        Subtotals(pos) = Prices(pos) * Requests(pos)
    Next pos
    
    Open App.Path & "\Customers\" & user & "\" & user & year & ".txt" For Output As #1
    
    For pos = 1 To counter
        Write #1, Items(pos), Prices(pos), Requests(pos), Subtotals(pos)
    Next pos
    
    Close #1
    
    'The following code prints all requested items under their respective
    'heading. Using this code in every rental form allows the display to always
    'be visible and up-to-date
    frmDisplay.picResults.Print Tab(40); "Prices"; Tab(50); "Number"; Tab(60); "Subtotals"
    frmDisplay.picResults.Print "--Motor Boats--"
    
    For pos = 1 To 9
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Canoes/Kayaks--"
   
    For pos = 10 To 27
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Rental Equipment--"
 
    For pos = 28 To 37
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Packs--"

    For pos = 38 To 42
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Tents--"
   
    For pos = 43 To 46
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Sleeping Bags/Pads--"

    For pos = 47 To 51
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Camp Items Misc.--"
 
    For pos = 52 To 64
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Transportation--"

    For pos = 65 To 86
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Towing--"

    For pos = 87 To 92
        If Requests(pos) > 0 Then
            frmDisplay.picResults.Print Items(pos); Tab(40); FormatCurrency(Prices(pos)); Tab(50); Requests(pos); Tab(60); FormatCurrency(Subtotals(pos))
        End If
    Next pos
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "--Guide Service--"

    If Requests(93) > 0 Then
        frmDisplay.picResults.Print Items(93); Tab(40); FormatCurrency(Prices(93)); Tab(50); Requests(93); Tab(60); FormatCurrency(Subtotals(93))
    End If
    
    'The following code first adds up the seperate subtotals of the rental itmes
    'and saves the values in certain groups so that they can be used for
    'calculating taxes. The subtotals and taxes are then added up along with a
    'certain group of items that are multiplies by the number of days in the
    'trip. The user ends up with a grand total for the trip.
    For pos = 1 To 64
        subtotal1 = subtotal1 + Subtotals(pos)
    Next pos
    
    For pos = 65 To counter
        subtotal2 = subtotal2 + Subtotals(pos)
    Next pos
    
    For pos = 87 To counter
        subtotal3 = subtotal3 + Subtotals(pos)
    Next pos
    
    For pos = 1 To 63
        grandTotal = grandTotal + (Subtotals(pos) * frmStartup.txtDays.Text)
    Next pos
    
    salesTax = subtotal1 * 0.065
    lodgingTax = Subtotals(64) * 0.03
    fsTax = subtotal3 * 0.03
    total = subtotal1 + subtotal2 + salesTax + lodgingTax + fsTax
    grandTotal = grandTotal + Subtotals(64) + subtotal2 + salesTax + lodgingTax + fsTax
    
    
    frmDisplay.picResults.Print
    frmDisplay.picResults.Print "Subtotal:"; Tab(25); FormatCurrency(subtotal1 + subtotal2)
    frmDisplay.picResults.Print "Sales Tax (6.5%):"; Tab(25); FormatCurrency(salesTax)
    frmDisplay.picResults.Print "Lodging Tax (3%):"; Tab(25); FormatCurrency(lodgingTax)
    frmDisplay.picResults.Print "USFS Tax (3%):"; Tab(25); FormatCurrency(fsTax)
    frmDisplay.picResults.Print "------------------------------------------------------"
    frmDisplay.picResults.Print "Total:"; Tab(25); FormatCurrency(total)
    If frmStartup.txtDays.Text > 0 Then
        frmDisplay.picResults.Print "For a " & frmStartup.txtDays.Text & " day trip, Grand Total ="; Tab(40); FormatCurrency(grandTotal)
    End If
End Sub

