VERSION 5.00
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   FillColor       =   &H0080FFFF&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   Palette         =   "FrmMain.frx":0000
   PaletteMode     =   2  'Custom
   Picture         =   "FrmMain.frx":35E0A
   ScaleHeight     =   4
   ScaleMode       =   5  'Inch
   ScaleWidth      =   5.833
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdaddnote 
      BackColor       =   &H000000FF&
      Caption         =   "Add Note"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdcurrentmonth 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to current month."
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdnextmonth 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5880
      Picture         =   "FrmMain.frx":78724
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdpreviousmonth 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      Picture         =   "FrmMain.frx":7B015
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdexit 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7920
      Picture         =   "FrmMain.frx":7D919
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdminimize 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7440
      Picture         =   "FrmMain.frx":80334
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Calender: By Chris Tift"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label lblnotes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   4920
      Width           =   8295
   End
   Begin VB.Label lblholiday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Holidays:"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   55
      Top             =   5280
      Width           =   8295
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   41
      Left            =   5760
      TabIndex        =   50
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   40
      Left            =   5160
      TabIndex        =   49
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   39
      Left            =   4560
      TabIndex        =   48
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   38
      Left            =   3960
      TabIndex        =   47
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   37
      Left            =   3360
      TabIndex        =   46
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   36
      Left            =   2760
      TabIndex        =   45
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   35
      Left            =   2040
      TabIndex        =   44
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   34
      Left            =   5760
      TabIndex        =   43
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   33
      Left            =   5160
      TabIndex        =   42
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   4560
      TabIndex        =   41
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   3960
      TabIndex        =   40
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   3360
      TabIndex        =   39
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   2760
      TabIndex        =   38
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   2040
      TabIndex        =   37
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   5760
      TabIndex        =   36
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   5160
      TabIndex        =   35
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   4560
      TabIndex        =   34
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   3960
      TabIndex        =   33
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   3360
      TabIndex        =   32
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   2760
      TabIndex        =   31
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   2040
      TabIndex        =   30
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   5760
      TabIndex        =   29
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   5160
      TabIndex        =   28
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   4560
      TabIndex        =   27
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   3960
      TabIndex        =   26
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   3360
      TabIndex        =   25
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   2760
      TabIndex        =   24
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   2040
      TabIndex        =   23
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   5760
      TabIndex        =   22
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   5160
      TabIndex        =   21
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   4560
      TabIndex        =   20
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   3960
      TabIndex        =   19
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   3360
      TabIndex        =   18
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   2760
      TabIndex        =   17
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2040
      TabIndex        =   16
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   5760
      TabIndex        =   15
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   5160
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4560
      TabIndex        =   13
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3960
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   11
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   1.417
      X2              =   4.333
      Y1              =   1.417
      Y2              =   1.417
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   630
      Left            =   3690
      TabIndex        =   7
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label lblcurrentdate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current:"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   5760
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   5160
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4560
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3960
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Calender
'Form name: FrmMain
'Author: Chris Tift
'Date Written: 10/30/05
'Objective: To create a working calander in which notes can be stored.

Private Sub cmdaddnote_Click()
notedate = InputBox("Enter the date you want the note on.", "Add Note")
addnote = InputBox("Enter Your Note:", "Add Note")
Days = 0
    Select Case TempMonth
        Case Is = 1
            Days = notedate
        Case Is > 1
            Days = 31
        Case Is > 2
            Days = 31 + 29
        Case Is > 3
            Days = 31 + 29 + 31
        Case Is > 4
            Days = 31 + 29 + 31 + 30
        Case Is > 5
            Days = 31 + 29 + 31 + 30 + 31
        Case Is > 6
            Days = 31 + 29 + 31 + 30 + 31 + 30
        Case Is > 7
            Days = 31 + 29 + 31 + 30 + 31 + 30 + 31
        Case Is > 8
            Days = 31 + 29 + 31 + 30 + 31 + 30 + 31 + 31
        Case Is > 9
            Days = 31 + 29 + 31 + 30 + 31 + 30 + 31 + 31 + 30
        Case Is > 10
            Days = 31 + 29 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31
        Case Is > 11
            Days = 31 + 29 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30
        Case Is = 12
            Days = 31 + 29 + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30
    End Select
    Days = Days + notedate
    Notes(Days) = addnote
    note1 = "Notes: " & notedate & ": " & addnote
lblnotes.Caption = note1
notedate = 0
addnote = 0

End Sub

Private Sub cmdcurrentmonth_Click()
TempMonth = CurrentMonth
TempYear = CurrentYear
monthstart = 6
 Select Case TempMonth    'determines what month you are looking at as well as what number in the array to start with and how many days are in that month
        Case Is = 1
        n = 1
        Days = 31
            lblMonth.ForeColor = &HFF&                      'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF&
        Next X
            lblcurrentdate.ForeColor = &HFF&
            lblnotes.ForeColor = &HFF&
            lblholiday.ForeColor = &HFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 2
    leapcheck = TempYear
    Fleap = False
    Do Until leapcheck < 2004
        leapcheck = leapcheck - 4
        If leapcheck = 2004 Then
            Fleap = True
        End If
    Loop
        n = 32
        Days = 29
            lblMonth.ForeColor = &H8080FF         'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H8080FF
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H8080FF
        Next X
            lblcurrentdate.ForeColor = &H8080FF
            lblnotes.ForeColor = &H8080FF
            lblholiday.ForeColor = &H8080FF
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 3
        n = 61
        Days = 31
            lblMonth.ForeColor = &HFFFF&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFFFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFFFF&
        Next X
            lblcurrentdate.ForeColor = &HFFFF&
            lblnotes.ForeColor = &HFFFF&
            lblholiday.ForeColor = &HFFFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 4
        n = 92
        Days = 30
            lblMonth.ForeColor = &H800000         'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H800000
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H800000
        Next X
            lblcurrentdate.ForeColor = &H800000
            lblnotes.ForeColor = &H800000
            lblholiday.ForeColor = &H800000
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 5
        n = 122
        Days = 31
            lblMonth.ForeColor = &HFF00&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF00&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF00&
        Next X
            lblcurrentdate.ForeColor = &HFF00&
            lblnotes.ForeColor = &HFF00&
            lblholiday.ForeColor = &HFF00&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 6
        n = 153
        Days = 30
            lblMonth.ForeColor = &HC0&                   'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0&
        Next X
            lblcurrentdate.ForeColor = &HC0&
            lblnotes.ForeColor = &HC0&
            lblholiday.ForeColor = &HC0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 7
        n = 183
        Days = 31
            lblMonth.ForeColor = &HC0C0C0                'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0C0C0
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0C0C0
        Next X
            lblcurrentdate.ForeColor = &HC0C0C0
            lblnotes.ForeColor = &HC0C0C0
            lblholiday.ForeColor = &HC0C0C0
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 8
        n = 214
        Days = 31
            lblMonth.ForeColor = &H40C0&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H40C0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H40C0&
        Next X
            lblcurrentdate.ForeColor = &H40C0&
            lblnotes.ForeColor = &H40C0&
            lblholiday.ForeColor = &H40C0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 9
        n = 244
        Days = 30
            lblMonth.ForeColor = &HFF&            'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF&
        Next X
            lblcurrentdate.ForeColor = &HFF&
            lblnotes.ForeColor = &HFF&
            lblholiday.ForeColor = &HFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 10
        n = 275
        Days = 31
            lblMonth.ForeColor = &H80FFFF                  'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H80FFFF
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H80FFFF
        Next X
            lblcurrentdate.ForeColor = &H80FFFF
            lblnotes.ForeColor = &H80FFFF
            lblholiday.ForeColor = &H80FFFF
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 11
        n = 305
        Days = 30
            lblMonth.ForeColor = &HFFFF00              'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFFFF00
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFFFF00
        Next X
            lblcurrentdate.ForeColor = &HFFFF00
            lblnotes.ForeColor = &HFFFF00
            lblholiday.ForeColor = &HFFFF00
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 12
        n = 336
        Days = 31
            lblMonth.ForeColor = &HC0&            'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0&
        Next X
            lblcurrentdate.ForeColor = &HC0&
            lblnotes.ForeColor = &HC0&
            lblholiday.ForeColor = &HC0&
            lbltitle.ForeColor = lblholiday.ForeColor
    End Select

     FrmMain.Picture = LoadPicture(App.Path & "\images\" & Monthw(TempMonth) & ".jpg")
       n = 1
    a = monthstart
        Do Until Month1(n) = TempMonth
            n = n + 1
        Loop
        Holi = "Holiday: "
        note1 = "Notes: "
     For X = n To n + Days
            a = a + 1
            If Notes(X) = "" Then
                GoTo H
            End If
            note1 = note1 + " " & Notes(X)
H:
            If Holiday(X) = "" Then
                GoTo F
            End If
            Holi = Holi + " " & Holiday(X)
F:
            If a < 0 Then
                a = 0
            End If
            lbldate(a).Caption = Date1(X)
                If a = monthstart + Days Then
                    monthend = 41 - (Days + monthstart + 1)
                     Do Until monthend <= 6
                        monthend = monthend - 6
                    Loop
                    monthstart = 5 - monthend
                    GoTo g
                End If
        Next X
g:
If Fleap = False And TempMonth = 2 Then   'if it isn't a leap year this clears out the 29 in februray
    lbldate(a - 1).Caption = ""
End If
   cmdcurrentmonth.Enabled = False
    lblholiday.Caption = Holi
    lblnotes.Caption = note1
    lblMonth.Caption = Monthw(TempMonth) & " of " & TempYear
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdminimize_Click()
FrmMain.WindowState = 1
End Sub

Private Sub cmdnextmonth_Click()
    For X = 0 To 41
        lbldate(X).Caption = ""
    Next X
    TempMonth = TempMonth + 1
    Select Case TempMonth
    Case Is = 1
        n = 1
        Days = 31
            lblMonth.ForeColor = &HFF&                      'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF&
        Next X
            lblcurrentdate.ForeColor = &HFF&
            lblnotes.ForeColor = &HFF&
            lblholiday.ForeColor = &HFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 2
    leapcheck = TempYear
    Fleap = False
    Do Until leapcheck < 2004
        leapcheck = leapcheck - 4
        If leapcheck = 2004 Then
            Fleap = True
        End If
    Loop
        n = 32
        Days = 29
            lblMonth.ForeColor = &H8080FF         'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H8080FF
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H8080FF
        Next X
            lblcurrentdate.ForeColor = &H8080FF
            lblnotes.ForeColor = &H8080FF
            lblholiday.ForeColor = &H8080FF
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 3
        n = 61
        Days = 31
            lblMonth.ForeColor = &HFFFF&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFFFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFFFF&
        Next X
            lblcurrentdate.ForeColor = &HFFFF&
            lblnotes.ForeColor = &HFFFF&
            lblholiday.ForeColor = &HFFFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 4
        n = 92
        Days = 30
            lblMonth.ForeColor = &H800000         'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H800000
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H800000
        Next X
            lblcurrentdate.ForeColor = &H800000
            lblnotes.ForeColor = &H800000
            lblholiday.ForeColor = &H800000
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 5
        n = 122
        Days = 31
            lblMonth.ForeColor = &HFF00&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF00&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF00&
        Next X
            lblcurrentdate.ForeColor = &HFF00&
            lblnotes.ForeColor = &HFF00&
            lblholiday.ForeColor = &HFF00&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 6
        n = 153
        Days = 30
            lblMonth.ForeColor = &HC0&                   'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0&
        Next X
            lblcurrentdate.ForeColor = &HC0&
            lblnotes.ForeColor = &HC0&
            lblholiday.ForeColor = &HC0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 7
        n = 183
        Days = 31
            lblMonth.ForeColor = &HC0C0C0                'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0C0C0
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0C0C0
        Next X
            lblcurrentdate.ForeColor = &HC0C0C0
            lblnotes.ForeColor = &HC0C0C0
            lblholiday.ForeColor = &HC0C0C0
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 8
        n = 214
        Days = 31
            lblMonth.ForeColor = &H40C0&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H40C0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H40C0&
        Next X
            lblcurrentdate.ForeColor = &H40C0&
            lblnotes.ForeColor = &H40C0&
            lblholiday.ForeColor = &H40C0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 9
        n = 244
        Days = 30
            lblMonth.ForeColor = &HFF&            'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF&
        Next X
            lblcurrentdate.ForeColor = &HFF&
            lblnotes.ForeColor = &HFF&
            lblholiday.ForeColor = &HFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 10
        n = 275
        Days = 31
            lblMonth.ForeColor = &H80FFFF                  'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H80FFFF
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H80FFFF
        Next X
            lblcurrentdate.ForeColor = &H80FFFF
            lblnotes.ForeColor = &H80FFFF
            lblholiday.ForeColor = &H80FFFF
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 11
        n = 305
        Days = 30
            lblMonth.ForeColor = &HFFFF00               'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFFFF00
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFFFF00
        Next X
            lblcurrentdate.ForeColor = &HFFFF00
            lblnotes.ForeColor = &HFFFF00
            lblholiday.ForeColor = &HFFFF00
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 12
        n = 336
        Days = 31
            lblMonth.ForeColor = &HC0&            'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0&
        Next X
            lblcurrentdate.ForeColor = &HC0&
            lblnotes.ForeColor = &HC0&
            lblholiday.ForeColor = &HC0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 13
        n = 1
        Days = 31
        TempMonth = 1
        TempYear = TempYear + 1
            lblMonth.ForeColor = &HFF&                      'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF&
        Next X
            lblcurrentdate.ForeColor = &HFF&
            lblnotes.ForeColor = &HFF&
            lblholiday.ForeColor = &HFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    End Select

     FrmMain.Picture = LoadPicture(App.Path & "\images\" & Monthw(TempMonth) & ".jpg")
    n = 1
    a = monthstart
        Do Until Month1(n) = TempMonth
            n = n + 1
        Loop
        Holi = "Holiday: "
        note1 = "Notes: "
    For X = n To n + Days
            a = a + 1
            If Notes(X) = "" Then
                GoTo H
            End If
            note1 = note1 + " " & Notes(X)
H:
            If Holiday(X) = "" Then
                GoTo F
            End If
            Holi = Holi + " " & Holiday(X)
F:
        If a < 0 Then
            a = 0
        End If
            lbldate(a).Caption = Date1(X)
            If a < 0 Then
                a = 0
            End If
            lbldate(a).Caption = Date1(X)
                If a = monthstart + Days Then
                    monthend = 41 - (Days + monthstart + 1)
                    Do Until monthend <= 7
                        monthend = monthend - 7
                    Loop
                    monthstart = 5 - monthend
                    GoTo g
                End If
        Next X
g:
    cmdcurrentmonth.Enabled = True
        If TempMonth = CurrentMonth Then
            cmdcurrentmonth.Enabled = False
        End If
    lblnotes.Caption = note1
    lblholiday.Caption = Holi
    lblMonth.Caption = Monthw(TempMonth) & " of " & TempYear
End Sub

Private Sub cmdpreviousmonth_Click()
    For X = 0 To 41
        lbldate(X).Caption = ""
    Next X
    TempMonth = TempMonth - 1
      Select Case TempMonth
       Case Is = 0
        n = 336
        Days = 31
        TempMonth = 12
        TempYear = TempYear - 1
            lblMonth.ForeColor = &HC0&            'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0&
        Next X
            lblcurrentdate.ForeColor = &HC0&
            lblnotes.ForeColor = &HC0&
            lblholiday.ForeColor = &HC0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 1
        n = 1
        Days = 31
            lblMonth.ForeColor = &HFF&                      'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF&
        Next X
            lblcurrentdate.ForeColor = &HFF&
            lblnotes.ForeColor = &HFF&
            lblholiday.ForeColor = &HFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 2
    leapcheck = TempYear
    Fleap = False
    Do Until leapcheck < 2004
        leapcheck = leapcheck - 4
        If leapcheck = 2004 Then
            Fleap = True
        End If
    Loop
        n = 32
        Days = 29
            lblMonth.ForeColor = &H8080FF         'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H8080FF
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H8080FF
        Next X
            lblcurrentdate.ForeColor = &H8080FF
            lblnotes.ForeColor = &H8080FF
            lblholiday.ForeColor = &H8080FF
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 3
        n = 61
        Days = 31
            lblMonth.ForeColor = &HFFFF&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFFFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFFFF&
        Next X
            lblcurrentdate.ForeColor = &HFFFF&
            lblnotes.ForeColor = &HFFFF&
            lblholiday.ForeColor = &HFFFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 4
        n = 92
        Days = 30
            lblMonth.ForeColor = &H800000         'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H800000
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H800000
        Next X
            lblcurrentdate.ForeColor = &H800000
            lblnotes.ForeColor = &H800000
            lblholiday.ForeColor = &H800000
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 5
        n = 122
        Days = 31
            lblMonth.ForeColor = &HFF00&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF00&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF00&
        Next X
            lblcurrentdate.ForeColor = &HFF00&
            lblnotes.ForeColor = &HFF00&
            lblholiday.ForeColor = &HFF00&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 6
        n = 153
        Days = 30
            lblMonth.ForeColor = &HC0&                   'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0&
        Next X
            lblcurrentdate.ForeColor = &HC0&
            lblnotes.ForeColor = &HC0&
            lblholiday.ForeColor = &HC0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 7
        n = 183
        Days = 31
            lblMonth.ForeColor = &HC0C0C0                'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0C0C0
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0C0C0
        Next X
            lblcurrentdate.ForeColor = &HC0C0C0
            lblnotes.ForeColor = &HC0C0C0
            lblholiday.ForeColor = &HC0C0C0
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 8
        n = 214
        Days = 31
            lblMonth.ForeColor = &H40C0&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H40C0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H40C0&
        Next X
            lblcurrentdate.ForeColor = &H40C0&
            lblnotes.ForeColor = &H40C0&
            lblholiday.ForeColor = &H40C0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 9
        n = 244
        Days = 30
            lblMonth.ForeColor = &HFF&            'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF&
        Next X
            lblcurrentdate.ForeColor = &HFF&
            lblnotes.ForeColor = &HFF&
            lblholiday.ForeColor = &HFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 10
        n = 275
        Days = 31
            lblMonth.ForeColor = &H80FFFF                  'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H80FFFF
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H80FFFF
        Next X
            lblcurrentdate.ForeColor = &H80FFFF
            lblnotes.ForeColor = &H80FFFF
            lblholiday.ForeColor = &H80FFFF
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 11
        n = 305
        Days = 30
            lblMonth.ForeColor = &HFFFF00              'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFFFF00
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFFFF00
        Next X
            lblcurrentdate.ForeColor = &HFFFF00
            lblnotes.ForeColor = &HFFFF00
            lblholiday.ForeColor = &HFFFF00
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 12
        n = 336
        Days = 31
            lblMonth.ForeColor = &HC0&            'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0&
        Next X
            lblcurrentdate.ForeColor = &HC0&
            lblnotes.ForeColor = &HC0&
            lblholiday.ForeColor = &HC0&
            lbltitle.ForeColor = lblholiday.ForeColor
    End Select

     FrmMain.Picture = LoadPicture(App.Path & "\images\" & Monthw(TempMonth) & ".jpg")
    n = 1
    a = monthstart
        Do Until Month1(n) = TempMonth
            n = n + 1
        Loop
        Holi = "Holiday: "
        note1 = "Notes: "
    For X = n To n + Days
            a = a + 1
            If Notes(X) = "" Then
                GoTo H
            End If
            note1 = note1 + " " & Notes(X)
H:
            If Holiday(X) = "" Then
                GoTo F
            End If
            Holi = Holi + " " & Holiday(X)
F:
            If a < 0 Then
                a = 0
            End If
            lbldate(a).Caption = Date1(X)
                If a = monthstart + Days Then
                    monthend = 41 - (Days + monthstart + 1)
                    Do Until monthend <= 7
                        monthend = monthend - 7
                    Loop
                    monthstart = 5 - monthend
                    GoTo g
                End If
        Next X
g:
    lblnotes.Caption = note1
    lblholiday.Caption = Holi
    lblMonth.Caption = Monthw(TempMonth) & " of " & TempYear
    cmdcurrentmonth.Enabled = True
        If TempMonth = CurrentMonth Then
        cmdcurrentmonth.Enabled = False
    End If
End Sub

Private Sub Form_Load()
Open App.Path & "\Dates.txt" For Input As #1
        For I = 1 To 366
            Input #1, Month1(I), Date1(I), Holiday(I), Notes(I)
        Next I
        Monthw(1) = "January"
        Monthw(2) = "February"
        Monthw(3) = "March"
        Monthw(4) = "April"
        Monthw(5) = "May"
        Monthw(6) = "June"
        Monthw(7) = "July"
        Monthw(8) = "August"
        Monthw(9) = "September"
        Monthw(10) = "October"
        Monthw(11) = "November"
        Monthw(12) = "December"
    lblcurrentdate.Caption = "Current: " & DateTime.Date
    CurrentMonth = Left(DateTime.Date, 2)   'gets the current month, day, year from computer clock
        TempMonth = CurrentMonth
    CurrentDay = Mid(DateTime.Date, 4, 2)
        Tempday = CurrentDay
    CurrentYear = Right(DateTime.Date, 4)
        TempYear = CurrentYear
    lblMonth.Caption = Monthw(CurrentMonth)
    FrmMain.Caption = DateTime.Date

    monthstart = 5
   Select Case TempMonth    'determines what month you are looking at as well as what number in the array to start with and how many days are in that month
    Case Is = 1
        n = 1
        Days = 31
            lblMonth.ForeColor = &HFF&                      'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF&
        Next X
            lblcurrentdate.ForeColor = &HFF&
            lblnotes.ForeColor = &HFF&
            lblholiday.ForeColor = &HFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 2
    leapcheck = TempYear
    Fleap = False
    Do Until leapcheck < 2004
        leapcheck = leapcheck - 4
        If leapcheck = 2004 Then
            Fleap = True
        End If
    Loop
        n = 32
        Days = 29
            lblMonth.ForeColor = &H8080FF         'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H8080FF
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H8080FF
        Next X
            lblcurrentdate.ForeColor = &H8080FF
            lblnotes.ForeColor = &H8080FF
            lblholiday.ForeColor = &H8080FF
            lbltitle.ForeColor = lblholiday.ForeColor
        Case Is = 3
        n = 61
        Days = 31
            lblMonth.ForeColor = &HFFFF&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFFFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFFFF&
        Next X
            lblcurrentdate.ForeColor = &HFFFF&
            lblnotes.ForeColor = &HFFFF&
            lblholiday.ForeColor = &HFFFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 4
        n = 92
        Days = 30
            lblMonth.ForeColor = &H800000         'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H800000
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H800000
        Next X
            lblcurrentdate.ForeColor = &H800000
            lblnotes.ForeColor = &H800000
            lblholiday.ForeColor = &H800000
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 5
        n = 122
        Days = 31
            lblMonth.ForeColor = &HFF00&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF00&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF00&
        Next X
            lblcurrentdate.ForeColor = &HFF00&
            lblnotes.ForeColor = &HFF00&
            lblholiday.ForeColor = &HFF00&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 6
        n = 153
        Days = 30
            lblMonth.ForeColor = &HC0&                   'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0&
        Next X
            lblcurrentdate.ForeColor = &HC0&
            lblnotes.ForeColor = &HC0&
            lblholiday.ForeColor = &HC0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 7
        n = 183
        Days = 31
            lblMonth.ForeColor = &HC0C0C0                'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0C0C0
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0C0C0
        Next X
            lblcurrentdate.ForeColor = &HC0C0C0
            lblnotes.ForeColor = &HC0C0C0
            lblholiday.ForeColor = &HC0C0C0
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 8
        n = 214
        Days = 31
            lblMonth.ForeColor = &H40C0&          'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H40C0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H40C0&
        Next X
            lblcurrentdate.ForeColor = &H40C0&
            lblnotes.ForeColor = &H40C0&
            lblholiday.ForeColor = &H40C0&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 9
        n = 244
        Days = 30
            lblMonth.ForeColor = &HFF&            'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFF&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFF&
        Next X
            lblcurrentdate.ForeColor = &HFF&
            lblnotes.ForeColor = &HFF&
            lblholiday.ForeColor = &HFF&
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 10
        n = 275
        Days = 31
            lblMonth.ForeColor = &H80FFFF                  'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &H80FFFF
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &H80FFFF
        Next X
            lblcurrentdate.ForeColor = &H80FFFF
            lblnotes.ForeColor = &H80FFFF
            lblholiday.ForeColor = &H80FFFF
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 11
        n = 305
        Days = 30
            lblMonth.ForeColor = &HFFFF00               'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HFFFF00
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HFFFF00
        Next X
            lblcurrentdate.ForeColor = &HFFFF00
            lblnotes.ForeColor = &HFFFF00
            lblholiday.ForeColor = &HFFFF00
            lbltitle.ForeColor = lblholiday.ForeColor
    Case Is = 12
        n = 336
        Days = 31
            lblMonth.ForeColor = &HC0&            'setting font color for each month, simply to see text better against background picture
        For X = 0 To 6
            lblDay(X).ForeColor = &HC0&
        Next X
        For X = 0 To 41
            lbldate(X).ForeColor = &HC0&
        Next X
            lblcurrentdate.ForeColor = &HC0&
            lblnotes.ForeColor = &HC0&
            lblholiday.ForeColor = &HC0&
            lbltitle.ForeColor = lblholiday.ForeColor
End Select
     FrmMain.Picture = LoadPicture(App.Path & "\images\" & Monthw(TempMonth) & ".jpg")
    a = monthstart
    Holi = "Holiday: "
    note1 = "Notes: "
         For X = n To n + Days   'writes the dates in the labels for the month
            a = a + 1
            If Notes(X) = "" Then
                GoTo H
            End If
            note1 = note1 + Date1(X) & ": " & Notes(X)
H:
            If Holiday(X) = "" Then
                GoTo F
            End If
            Holi = Holi + " " & Holiday(X)
F:
            lbldate(a).Caption = Date1(X)
                If a = monthstart + Days Then
                    monthend = 41 - (Days + monthstart + 1)
                     Do Until monthend <= 6
                        monthend = monthend - 6
                    Loop
                    monthstart = 5 - monthend
                    GoTo g
                End If
        Next X
g:
If Fleap = False And TempMonth = 2 Then   'if it isn't a leap year this clears out the 29 in februray
    lbldate(a).Caption = ""
End If
    lblnotes.Caption = note1
    lblholiday.Caption = Holi
    lblMonth.Caption = Monthw(TempMonth) & " of " & TempYear
    
  
End Sub


