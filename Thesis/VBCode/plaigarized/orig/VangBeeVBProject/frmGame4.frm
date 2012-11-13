VERSION 5.00
Begin VB.Form frmGame4 
   Caption         =   "President Sorting Game"
   ClientHeight    =   11520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   Picture         =   "frmGame4.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   8280
      TabIndex        =   94
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   12840
      TabIndex        =   92
      Top             =   9480
      Width           =   2295
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Index"
      Height          =   615
      Left            =   10320
      TabIndex        =   91
      Top             =   9480
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   8775
      Left            =   10320
      ScaleHeight     =   8715
      ScaleWidth      =   8235
      TabIndex        =   90
      Top             =   600
      Width           =   8295
   End
   Begin VB.CommandButton cmdSorting 
      Caption         =   "Sort all the president in numerical order"
      Height          =   615
      Left            =   8280
      TabIndex        =   89
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdCheckAnswer 
      Caption         =   "Check My Answer"
      Height          =   735
      Left            =   360
      TabIndex        =   44
      Top             =   8640
      Width           =   6135
   End
   Begin VB.TextBox txtPres37 
      Height          =   285
      Index           =   43
      Left            =   3360
      TabIndex        =   88
      Top             =   8160
      Width           =   375
   End
   Begin VB.TextBox txtPres5 
      Height          =   285
      Index           =   42
      Left            =   3360
      TabIndex        =   87
      Top             =   7800
      Width           =   375
   End
   Begin VB.TextBox txtPres34 
      Height          =   285
      Index           =   41
      Left            =   3360
      TabIndex        =   86
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox txtPres10 
      Height          =   285
      Index           =   40
      Left            =   3360
      TabIndex        =   85
      Top             =   7080
      Width           =   375
   End
   Begin VB.TextBox txtPres13 
      Height          =   285
      Index           =   39
      Left            =   3360
      TabIndex        =   84
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtPres31 
      Height          =   285
      Index           =   38
      Left            =   3360
      TabIndex        =   83
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox txtPres42 
      Height          =   285
      Index           =   37
      Left            =   3360
      TabIndex        =   82
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox txtPres29 
      Height          =   285
      Index           =   36
      Left            =   3360
      TabIndex        =   81
      Top             =   5640
      Width           =   375
   End
   Begin VB.TextBox txtPres27 
      Height          =   285
      Index           =   35
      Left            =   3360
      TabIndex        =   80
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtPres26 
      Height          =   285
      Index           =   34
      Left            =   3360
      TabIndex        =   79
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox txtPres24 
      Height          =   285
      Index           =   33
      Left            =   3360
      TabIndex        =   78
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox txtPres23 
      Height          =   285
      Index           =   32
      Left            =   3360
      TabIndex        =   77
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txtPres19 
      Height          =   285
      Index           =   31
      Left            =   3360
      TabIndex        =   76
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox txtPres36 
      Height          =   285
      Index           =   30
      Left            =   3360
      TabIndex        =   75
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox txtPres44 
      Height          =   285
      Index           =   29
      Left            =   3360
      TabIndex        =   74
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtPres22 
      Height          =   285
      Index           =   28
      Left            =   3360
      TabIndex        =   73
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtPres39 
      Height          =   285
      Index           =   27
      Left            =   3360
      TabIndex        =   72
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtPres20 
      Height          =   285
      Index           =   26
      Left            =   3360
      TabIndex        =   71
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox txtPres3 
      Height          =   285
      Index           =   25
      Left            =   3360
      TabIndex        =   70
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtPres33 
      Height          =   285
      Index           =   24
      Left            =   3360
      TabIndex        =   69
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtPres18 
      Height          =   285
      Index           =   23
      Left            =   3360
      TabIndex        =   68
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtPres25 
      Height          =   285
      Index           =   22
      Left            =   3360
      TabIndex        =   67
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtPres17 
      Height          =   285
      Index           =   21
      Left            =   360
      TabIndex        =   66
      Top             =   8160
      Width           =   375
   End
   Begin VB.TextBox txtPres41 
      Height          =   285
      Index           =   20
      Left            =   360
      TabIndex        =   65
      Top             =   7800
      Width           =   375
   End
   Begin VB.TextBox txtPres16 
      Height          =   285
      Index           =   19
      Left            =   360
      TabIndex        =   64
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox txtPres35 
      Height          =   285
      Index           =   18
      Left            =   360
      TabIndex        =   63
      Top             =   7080
      Width           =   375
   End
   Begin VB.TextBox txtPres1 
      Height          =   285
      Index           =   17
      Left            =   360
      TabIndex        =   62
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtPres15 
      Height          =   285
      Index           =   16
      Left            =   360
      TabIndex        =   61
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox txtPres30 
      Height          =   285
      Index           =   15
      Left            =   360
      TabIndex        =   60
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox txtPres14 
      Height          =   285
      Index           =   14
      Left            =   360
      TabIndex        =   59
      Top             =   5640
      Width           =   375
   End
   Begin VB.TextBox txtPres32 
      Height          =   285
      Index           =   13
      Left            =   360
      TabIndex        =   58
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtPres12 
      Height          =   285
      Index           =   12
      Left            =   360
      TabIndex        =   57
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox txtPres21 
      Height          =   285
      Index           =   11
      Left            =   360
      TabIndex        =   56
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox txtPres11 
      Height          =   285
      Index           =   10
      Left            =   360
      TabIndex        =   55
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txtPres9 
      Height          =   285
      Index           =   9
      Left            =   360
      TabIndex        =   54
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox txtPres28 
      Height          =   285
      Index           =   8
      Left            =   360
      TabIndex        =   53
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox txtPres7 
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   52
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtPres8 
      Height          =   285
      Index           =   6
      Left            =   360
      TabIndex        =   51
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtPres6 
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   50
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtPres40 
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   49
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox txtPres4 
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   48
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtPres43 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   47
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtPres2 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   46
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtPres38 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   45
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblDefinition 
      Caption         =   $"frmGame4.frx":29500
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   3015
      Left            =   6720
      TabIndex        =   93
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblPres37 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Richard M. Nixon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   43
      Top             =   8160
      Width           =   2535
   End
   Begin VB.Label lblPres5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "James Monroe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   42
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label lblPres34 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dwight D. Eisenhower"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   41
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label lblPres10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "John Tyler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   40
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label lblPres13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Millard Fillmore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   39
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label lblPres31 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Herbert Hoover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   38
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label lblPres42 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bill Clinton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   37
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label lblPres29 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Warren G. Harding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   36
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label lblPres27 
      BackColor       =   &H00FFFFFF&
      Caption         =   "William Howard Taft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   35
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label lblPres26 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Theodore Roosevelt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   34
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label lblPres24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grover Cleveland (2nd term)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   33
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblPres23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Benjamin Harrison"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   32
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblPres19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rutherford B. Hayes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   31
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lblPres36 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lyndon B. Johnson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   30
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblPres44 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Barack Obama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   29
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblPres22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grover Cleveland"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   28
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label lblPres39 
      BackColor       =   &H00FFFFFF&
      Caption         =   "James Carter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   27
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblPres20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "James Garfield"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   26
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblpres33 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Harry S. Truman"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblPres3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thomas Jefferson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblPres18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ulysses S. Grant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   23
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblPres25 
      BackColor       =   &H00FFFFFF&
      Caption         =   "William McKinley"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblPres17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Andrew Johnson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Label lblPres41 
      BackColor       =   &H00FFFFFF&
      Caption         =   "George H. W. Bush"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblPres16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Abraham Lincoln"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label lblPres35 
      BackColor       =   &H00FFFFFF&
      Caption         =   "John F. Kennedy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lblPres1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "George Washington"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label lblPres15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "James Buchanon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label lblPres30 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calvin Coolidge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label lblPres14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Franklin Pierce"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label lblPres32 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Franklin D. Roosevelt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label lblPres12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zachary Taylor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblPres21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chester A. Arthur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblPres11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "James K. Polk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblPres9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "William Henry Harrison"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblPres28 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Woodrow Wilson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblPres7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Andrew Jackson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblPres8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Martin Van Buren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblPres6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "John Quincy Adams"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lblPres40 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ronald Reagon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblPres4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "James Madison"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblPres43 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Geroge W. Bush"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblPres2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "John Adams"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblPres38 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gerald R. Ford"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "frmGame4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim presidentnumber(1 To 50) As Integer
Dim presidentname(1 To 50) As String
Dim ctr As Integer, pos As Integer, pass As Integer



Private Sub cmdBack_Click()

frmGameScene.Show
frmGame4.Hide

End Sub

Private Sub cmdCheckAnswer_Click()

Dim presidentnumber(1 To 50) As String
Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x, y, z As String
Dim aa, bb, cc, dd, ee, ff, gg, hh, ii, jj, kk, ll, mm, nn, oo, pp, qq, rr As String
Dim ctr As Integer, pos As Integer, pass As Integer

a = txtPres38(0).Text
b = txtPres2(1).Text
c = txtPres43(2).Text
d = txtPres4(3).Text
e = txtPres40(4).Text
f = txtPres6(5).Text
g = txtPres8(6).Text
h = txtPres7(7).Text
i = txtPres28(8).Text
j = txtPres9(9).Text
k = txtPres11(10).Text
l = txtPres21(11).Text
m = txtPres12(12).Text
n = txtPres32(13).Text
o = txtPres14(14).Text
p = txtPres30(15).Text
q = txtPres15(16).Text
r = txtPres1(17).Text
s = txtPres35(18).Text
t = txtPres16(19).Text
u = txtPres41(20).Text
v = txtPres17(21).Text
w = txtPres25(22).Text
x = txtPres18(23).Text
y = txtPres33(24).Text
z = txtPres3(25).Text
aa = txtPres20(26).Text
bb = txtPres39(27).Text
cc = txtPres22(28).Text
dd = txtPres44(29).Text
ee = txtPres36(30).Text
ff = txtPres19(31).Text
gg = txtPres23(32).Text
hh = txtPres24(33).Text
ii = txtPres26(34).Text
jj = txtPres27(35).Text
kk = txtPres29(36).Text
ll = txtPres42(37).Text
mm = txtPres31(38).Text
nn = txtPres13(39).Text
oo = txtPres10(40).Text
pp = txtPres34(41).Text
qq = txtPres5(42).Text
rr = txtPres37(43).Text

Open App.Path & "\presidents.txt" For Input As #1

Do While Not EOF(1)
    pos = pos + 1
    Input #1, presidentnumber(pos), presidentname(pos)
Loop
    
Close #1

    If a = presidentnumber(1) Then
        ctr = ctr + 1
    End If
    
    If b = presidentnumber(2) Then
        ctr = ctr + 1
    End If
    
    If c = presidentnumber(3) Then
        ctr = ctr + 1
    End If
    
    If d = presidentnumber(4) Then
        ctr = ctr + 1
    End If
    
    If e = presidentnumber(5) Then
        ctr = ctr + 1
    End If
    
    If f = presidentnumber(6) Then
        ctr = ctr + 1
    End If
    
    If g = presidentnumber(7) Then
        ctr = ctr + 1
    End If
    
    If h = presidentnumber(8) Then
        ctr = ctr + 1
    End If
    
    If i = presidentnumber(9) Then
        ctr = ctr + 1
    End If
    
    If j = presidentnumber(10) Then
        ctr = ctr + 1
    End If
    
    If k = presidentnumber(11) Then
        ctr = ctr + 1
    End If
    
    If l = presidentnumber(12) Then
        ctr = ctr + 1
    End If
    
    If m = presidentnumber(13) Then
        ctr = ctr + 1
    End If
    
    If n = presidentnumber(14) Then
        ctr = ctr + 1
    End If
    
    If o = presidentnumber(15) Then
        ctr = ctr + 1
    End If
    
    If p = presidentnumber(16) Then
        ctr = ctr + 1
    End If
    
    If q = presidentnumber(17) Then
        ctr = ctr + 1
    End If
    
    If r = presidentnumber(18) Then
        ctr = ctr + 1
    End If
    
    If s = presidentnumber(19) Then
        ctr = ctr + 1
    End If
    
    If t = presidentnumber(20) Then
        ctr = ctr + 1
    End If
    
    If u = presidentnumber(21) Then
        ctr = ctr + 1
    End If
    
    If v = presidentnumber(22) Then
        ctr = ctr + 1
    End If
    
    If w = presidentnumber(23) Then
        ctr = ctr + 1
    End If
    
    If x = presidentnumber(24) Then
        ctr = ctr + 1
    End If
    
    If y = presidentnumber(25) Then
        ctr = ctr + 1
    End If
    
    If z = presidentnumber(26) Then
        ctr = ctr + 1
    End If
    
    If aa = presidentnumber(27) Then
        ctr = ctr + 1
    End If
    
    If bb = presidentnumber(28) Then
        ctr = ctr + 1
    End If
    
    If cc = presidentnumber(29) Then
        ctr = ctr + 1
    End If
    
    If dd = presidentnumber(30) Then
        ctr = ctr + 1
    End If
    
    If ee = presidentnumber(31) Then
        ctr = ctr + 1
    End If
    
    If ff = presidentnumber(32) Then
        ctr = ctr + 1
    End If
    
    If gg = presidentnumber(33) Then
        ctr = ctr + 1
    End If
    
    If hh = presidentnumber(34) Then
        ctr = ctr + 1
    End If
    
    If ii = presidentnumber(35) Then
        ctr = ctr + 1
    End If
    
    If jj = presidentnumber(36) Then
        ctr = ctr + 1
    End If
    
    If kk = presidentnumber(37) Then
        ctr = ctr + 1
    End If
    
    If ll = presidentnumber(38) Then
        ctr = ctr + 1
    End If
    
    If mm = presidentnumber(39) Then
        ctr = ctr + 1
    End If
    
    If nn = presidentnumber(40) Then
        ctr = ctr + 1
    End If
    
    If oo = presidentnumber(41) Then
        ctr = ctr + 1
    End If
    
    If pp = presidentnumber(42) Then
        ctr = ctr + 1
    End If
    
    If qq = presidentnumber(43) Then
        ctr = ctr + 1
    End If
    
    If rr = presidentnumber(44) Then
        ctr = ctr + 1
    End If
    
MsgBox "You got " & ctr & " correct!"

End Sub

Private Sub cmdClear_Click()

picResults.Cls

End Sub

Private Sub cmdQuit_Click()

End

End Sub

Private Sub cmdSorting_Click()

Dim temppresidentnumber As Integer
Dim temppresidentname As String

ctr = 0
pos = 0

Open App.Path & "\presidents.txt" For Input As #2
ctr = 0
Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, presidentnumber(ctr), presidentname(ctr)
Loop
    
Close #2

For pass = 1 To ctr
    For pos = 1 To ctr - pass
        If presidentnumber(pos) > presidentnumber(pos + 1) Then
            temppresidentnumber = presidentnumber(pos)
            presidentnumber(pos) = presidentnumber(pos + 1)
            presidentnumber(pos + 1) = temppresidentnumber
                
            temppresidentname = presidentname(pos)
            presidentname(pos) = presidentname(pos + 1)
            presidentname(pos + 1) = temppresidentname
                
        End If
    Next pos
Next pass

For pos = 1 To ctr
    picResults.Print presidentnumber(pos), presidentname(pos)
Next pos
pos = pos + 1

End Sub
