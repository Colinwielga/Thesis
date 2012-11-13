VERSION 5.00
Begin VB.Form frmcanoes 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Wilderness Outfitters"
   ClientHeight    =   10200
   ClientLeft      =   1065
   ClientTop       =   945
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmcanoes.frx":0000
   ScaleHeight     =   10200
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdPreviousCanoes 
      Caption         =   "Previous Page"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   62
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton cmdSubmitCanoes 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   61
      Top             =   9600
      Width           =   1695
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   240
      TabIndex        =   19
      Text            =   "0"
      Top             =   9120
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   240
      TabIndex        =   18
      Text            =   "0"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   240
      TabIndex        =   17
      Text            =   "0"
      Top             =   7920
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   240
      TabIndex        =   16
      Text            =   "0"
      Top             =   7560
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   240
      TabIndex        =   15
      Text            =   "0"
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   14
      Text            =   "0"
      Top             =   6360
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   13
      Text            =   "0"
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   12
      Text            =   "0"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   10
      Text            =   "0"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   9
      Text            =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Text            =   "0"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Text            =   "0"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Text            =   "0"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Text            =   "0"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Text            =   "0"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtCanoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Text            =   "0"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label48 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "No."
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
      Left            =   240
      TabIndex        =   67
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label47 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "No."
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
      Left            =   240
      TabIndex        =   66
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label46 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "No."
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
      Left            =   240
      TabIndex        =   65
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label45 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "No."
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
      Left            =   240
      TabIndex        =   64
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "No."
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
      Left            =   240
      TabIndex        =   63
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label44 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$28.00"
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
      Left            =   6120
      TabIndex        =   60
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label43 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$28.00"
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
      Left            =   6120
      TabIndex        =   59
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Label Label42 
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
      Left            =   6120
      TabIndex        =   58
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label41 
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
      Left            =   6120
      TabIndex        =   57
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label Label40 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$28.00"
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
      Left            =   6120
      TabIndex        =   56
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label39 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$42.00"
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
      Left            =   6120
      TabIndex        =   55
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label38 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$42.00"
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
      Left            =   6120
      TabIndex        =   54
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label37 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$42.00"
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
      Left            =   6120
      TabIndex        =   53
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$42.00"
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
      Left            =   6120
      TabIndex        =   52
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label35 
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
      Left            =   6120
      TabIndex        =   51
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label34 
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
      Left            =   6120
      TabIndex        =   50
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label33 
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
      Left            =   6120
      TabIndex        =   49
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label32 
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
      Left            =   6120
      TabIndex        =   48
      Top             =   3360
      Width           =   1335
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
      Left            =   6120
      TabIndex        =   47
      Top             =   3000
      Width           =   1335
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
      Left            =   6120
      TabIndex        =   46
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$32.00"
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
      Left            =   6120
      TabIndex        =   45
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label28 
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
      Left            =   6120
      TabIndex        =   44
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$22.00"
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
      Left            =   6120
      TabIndex        =   43
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Unit Rate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   42
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "17' Current Design Storm (64 lbs.)"
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
      Left            =   960
      TabIndex        =   41
      Top             =   9120
      Width           =   3975
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "16.5' Current Design Squall (59 lbs.)"
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
      Left            =   960
      TabIndex        =   40
      Top             =   8760
      Width           =   3975
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "17' We-no-nah Encounter (38 lbs.)"
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
      Left            =   960
      TabIndex        =   39
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "16' Souris River Quetico Kevlar (42 lbs.)"
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
      Left            =   960
      TabIndex        =   38
      Top             =   7560
      Width           =   4215
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "15.5' Souris River Tranquility Duralite (37 lbs.)"
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
      Left            =   960
      TabIndex        =   37
      Top             =   7200
      Width           =   4215
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "20' We-no-nah MN III (55 lbs., 3 seats)"
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
      Left            =   960
      TabIndex        =   36
      Top             =   6360
      Width           =   4215
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "19' We-no-nah Seneca (55 lbs., 3 seats)"
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
      Left            =   960
      TabIndex        =   35
      Top             =   6000
      Width           =   4215
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "19' We-no-nah Itasca Kevlar (52 lbs.)"
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
      Left            =   960
      TabIndex        =   34
      Top             =   5640
      Width           =   4335
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "18.5' Souris River Quetico Kevlar (48 lbs., 3 seats)"
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
      Left            =   960
      TabIndex        =   33
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "18.5' Alumacraft (69 lbs.)"
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
      Left            =   960
      TabIndex        =   32
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "18' We-no-nah Champlain (45 lbs.)"
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
      Left            =   960
      TabIndex        =   31
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "18.5' We-no-nah MN II (44 lbs.)"
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
      Left            =   960
      TabIndex        =   30
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "18.5' Souris River Quetico Kevlar (48 lbs.)"
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
      Left            =   960
      TabIndex        =   29
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "17' Souris River Quetico (44 lbs.)"
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
      Left            =   960
      TabIndex        =   28
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "17' We-no-nah Boundary Waters (42 lbs.)"
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
      Left            =   960
      TabIndex        =   27
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "17' Souris River Quetico Duralite (54 lbs.)"
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
      Left            =   960
      TabIndex        =   26
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "16' Souris River Quetico Duralite (49 lbs.)"
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
      Left            =   960
      TabIndex        =   25
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "17' Lightweight aluminum (65 lbs.)"
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
      Left            =   960
      TabIndex        =   24
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Kayaks"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   23
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Solo Canoes"
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
      Left            =   960
      TabIndex        =   22
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "3-Person Canoes"
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
      Left            =   960
      TabIndex        =   21
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tandem Kevlars"
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
      Left            =   960
      TabIndex        =   20
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tandem Canoes"
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
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Canoes and Kayaks"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmcanoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wilderness Outfitters Partial Outfitting
'Justin Bork
'March, 2009
'
'frmCanoes
'The purpose of this form is to allow the user to choose from the Canoes and
'kayaks available for rental.

Private Sub Command2_Click()
End Sub
Private Sub Command1_Click()
End Sub

Private Sub cmdPreviousCanoes_Click()
    frmcanoes.Hide
    frmStartup.Show
End Sub

Private Sub cmdSubmitCanoes_Click()
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
    
    For pos = 10 To 27
        i = i + 1
        Requests(pos) = txtCanoes(i).Text
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
