VERSION 5.00
Begin VB.Form frmSudoku 
   Caption         =   "Sudoku"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picB3 
      Height          =   375
      Left            =   1200
      Picture         =   "frmSudoku.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   84
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4920
      TabIndex        =   82
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdBackToMain 
      Caption         =   "Go Back To Main Screen"
      Height          =   735
      Left            =   4920
      TabIndex        =   81
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdCheckAnswers 
      Caption         =   "Check Answers!"
      Height          =   735
      Left            =   4920
      TabIndex        =   80
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtI9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   79
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picH9 
      Height          =   375
      Left            =   3840
      Picture         =   "frmSudoku.frx":0582
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   78
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtG9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   77
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtI8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   76
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox picH8 
      Height          =   375
      Left            =   3840
      Picture         =   "frmSudoku.frx":0B04
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   75
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox picG8 
      Height          =   375
      Left            =   3480
      Picture         =   "frmSudoku.frx":1086
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   74
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox picI7 
      Height          =   375
      Left            =   4200
      Picture         =   "frmSudoku.frx":1608
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   73
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox picH7 
      Height          =   375
      Left            =   3840
      Picture         =   "frmSudoku.frx":1B8A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   72
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtG7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   71
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtF9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   70
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picF8 
      Height          =   375
      Left            =   2880
      Picture         =   "frmSudoku.frx":210C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   69
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox picF7 
      Height          =   375
      Left            =   2880
      Picture         =   "frmSudoku.frx":268E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   68
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox picE9 
      Height          =   375
      Left            =   2520
      Picture         =   "frmSudoku.frx":2C10
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   67
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picE7 
      Height          =   375
      Left            =   2520
      Picture         =   "frmSudoku.frx":3192
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   66
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtE8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   65
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtD9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   64
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtD8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   63
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtD7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   62
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox picC9 
      Height          =   375
      Left            =   1560
      Picture         =   "frmSudoku.frx":3714
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   61
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picC8 
      Height          =   375
      Left            =   1560
      Picture         =   "frmSudoku.frx":3C96
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   60
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtB9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   59
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtB8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   58
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox picA9 
      Height          =   375
      Left            =   840
      Picture         =   "frmSudoku.frx":4218
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   57
      Top             =   4440
      Width           =   375
   End
   Begin VB.PictureBox picA8 
      Height          =   375
      Left            =   840
      Picture         =   "frmSudoku.frx":479A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   56
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtC7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   55
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtB7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   54
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtA7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   53
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox picI6 
      Height          =   375
      Left            =   4200
      Picture         =   "frmSudoku.frx":4D1C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   52
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox picH6 
      Height          =   375
      Left            =   3840
      Picture         =   "frmSudoku.frx":529E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   51
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtI5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   50
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtH5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   49
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtI4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   48
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtH4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   47
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtG6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   46
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox picG5 
      Height          =   375
      Left            =   3480
      Picture         =   "frmSudoku.frx":5820
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   45
      Top             =   2760
      Width           =   375
   End
   Begin VB.PictureBox picG4 
      Height          =   375
      Left            =   3480
      Picture         =   "frmSudoku.frx":5DA2
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   44
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtD6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   43
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtE6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   42
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox picF6 
      Height          =   375
      Left            =   2880
      Picture         =   "frmSudoku.frx":6324
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   41
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtF5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   40
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtE5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   39
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtD5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   38
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtF4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   37
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtE4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   36
      Top             =   2400
      Width           =   375
   End
   Begin VB.PictureBox picD4 
      Height          =   375
      Left            =   2160
      Picture         =   "frmSudoku.frx":68A6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   35
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtA6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   34
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtB6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   33
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtA5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   32
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtB5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   31
      Top             =   2760
      Width           =   375
   End
   Begin VB.PictureBox picC6 
      Height          =   375
      Left            =   1560
      Picture         =   "frmSudoku.frx":6E28
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   30
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox picC5 
      Height          =   375
      Left            =   1560
      Picture         =   "frmSudoku.frx":73AA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   29
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtC4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   28
      Top             =   2400
      Width           =   375
   End
   Begin VB.PictureBox picB4 
      Height          =   375
      Left            =   1200
      Picture         =   "frmSudoku.frx":792C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   27
      Top             =   2400
      Width           =   375
   End
   Begin VB.PictureBox picA4 
      Height          =   375
      Left            =   840
      Picture         =   "frmSudoku.frx":7EAE
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   26
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtI3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtH3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   1800
      Width           =   375
   End
   Begin VB.PictureBox picI2 
      Height          =   375
      Left            =   4200
      Picture         =   "frmSudoku.frx":8430
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   23
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtH2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   22
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox picI1 
      Height          =   375
      Left            =   4200
      Picture         =   "frmSudoku.frx":89B2
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   21
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtH1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtG3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   1800
      Width           =   375
   End
   Begin VB.PictureBox picG2 
      Height          =   375
      Left            =   3480
      Picture         =   "frmSudoku.frx":8F34
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox picG1 
      Height          =   375
      Left            =   3480
      Picture         =   "frmSudoku.frx":94B6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtF3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtF2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtF1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox picE3 
      Height          =   375
      Left            =   2520
      Picture         =   "frmSudoku.frx":9A38
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtE2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox picD3 
      Height          =   375
      Left            =   2160
      Picture         =   "frmSudoku.frx":9FBA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   1800
      Width           =   375
   End
   Begin VB.PictureBox picD2 
      Height          =   375
      Left            =   2160
      Picture         =   "frmSudoku.frx":A53C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox picE1 
      Height          =   375
      Left            =   2520
      Picture         =   "frmSudoku.frx":AABE
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtD1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtA2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtC1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtC3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtA1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox picA3 
      Height          =   375
      Left            =   840
      Picture         =   "frmSudoku.frx":B040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   1800
      Width           =   375
   End
   Begin VB.PictureBox picB2 
      Height          =   375
      Left            =   1200
      Picture         =   "frmSudoku.frx":B5C2
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox picC2 
      Height          =   375
      Left            =   1560
      Picture         =   "frmSudoku.frx":BB44
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox picB1 
      Height          =   375
      Left            =   1200
      Picture         =   "frmSudoku.frx":C0C6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblSudoku 
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   83
      Top             =   240
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   840
      X2              =   4560
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   3360
      X2              =   3360
      Y1              =   1080
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   840
      X2              =   4440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   1080
      Y2              =   4800
   End
End
Attribute VB_Name = "frmSudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBackToMain_Click()
frmSudoku.Hide
frmMainScreen.Show
End Sub

Private Sub cmdCheckAnswers_Click()
Dim Solved As Boolean
Dim CTR As Integer
Solved = True
If txtA1.Text = "" Or txtA1.Text <> "3" Then
    CTR = CTR + 1
    Solved = False
End If
If txtA2.Text = "" Or txtA2.Text <> "6" Then
    CTR = CTR + 1
    Solved = False
End If
If txtA5.Text = "" Or txtA5.Text <> "2" Then
    CTR = CTR + 1
    Solved = False
End If
If txtA6.Text = "" Or txtA6.Text <> "8" Then
    CTR = CTR + 1
    Solved = False
End If
If txtA7.Text = "" Or txtA7.Text <> "7" Then
    CTR = CTR + 1
    Solved = False
End If
If txtB5.Text = "" Or txtB5.Text <> "6" Then
    CTR = CTR + 1
    Solved = False
End If
If txtB6.Text = "" Or txtB6.Text <> "4" Then
    CTR = CTR + 1
    Solved = False
End If
If txtB7.Text = "" Or txtB7.Text <> "3" Then
    CTR = CTR + 1
    Solved = False
End If
If txtB8.Text = "" Or txtB8.Text <> "8" Then
    CTR = CTR + 1
    Solved = False
End If
If txtC1.Text = "" Or txtC1.Text <> "4" Then
    CTR = CTR + 1
    Solved = False
End If
If txtC3.Text = "" Or txtC3.Text <> "8" Then
    CTR = CTR + 1
    Solved = False
End If
If txtC4.Text = "" Or txtC4.Text <> "3" Then
    CTR = CTR + 1
    Solved = False
End If
If txtC7.Text = "" Or txtC7.Text <> "2" Then
    CTR = CTR + 1
    Solved = False
End If
If txtD1.Text = "" Or txtD1.Text <> "6" Then
    CTR = CTR + 1
    Solved = False
End If
If txtD5.Text = "" Or txtD5.Text <> "8" Then
    CTR = CTR + 1
    Solved = False
End If
If txtD6.Text = "" Or txtD6.Text <> "9" Then
    CTR = CTR + 1
    Solved = False
End If
If txtD7.Text = "" Or txtD7.Text <> "5" Then
    CTR = CTR + 1
    Solved = False
End If
If txtD8.Text = "" Or txtD8.Text <> "4" Then
    CTR = CTR + 1
    Solved = False
End If
If txtD9.Text = "" Or txtD9.Text <> "7" Then
    CTR = CTR + 1
    Solved = False
End If
If txtE2.Text = "" Or txtE2.Text <> "8" Then
    CTR = CTR + 1
    Solved = False
End If
If txtE4.Text = "" Or txtE4.Text <> "4" Then
    CTR = CTR + 1
    Solved = False
End If
If txtE5.Text = "" Or txtE5.Text <> "3" Then
    CTR = CTR + 1
    Solved = False
End If
If txtE6.Text = "" Or txtE6.Text <> "5" Then
    CTR = CTR + 1
    Solved = False
End If
If txtE8.Text = "" Or txtE8.Text <> "6" Then
    CTR = CTR + 1
    Solved = False
End If
If txtF1.Text = "" Or txtF1.Text <> "9" Then
    CTR = CTR + 1
    Solved = False
End If
If txtF2.Text = "" Or txtF2.Text <> "5" Then
    CTR = CTR + 1
    Solved = False
End If
If txtF3.Text = "" Or txtF3.Text <> "4" Then
    CTR = CTR + 1
    Solved = False
End If
If txtF4.Text = "" Or txtF4.Text <> "1" Then
    CTR = CTR + 1
    Solved = False
End If
If txtF5.Text = "" Or txtF5.Text <> "7" Then
    CTR = CTR + 1
    Solved = False
End If
If txtF9.Text = "" Or txtF9.Text <> "2" Then
    CTR = CTR + 1
    Solved = False
End If
If txtG3.Text = "" Or txtG3.Text <> "7" Then
    CTR = CTR + 1
    Solved = False
End If
If txtG6.Text = "" Or txtG6.Text <> "2" Then
    CTR = CTR + 1
    Solved = False
End If
If txtG7.Text = "" Or txtG7.Text <> "4" Then
    CTR = CTR + 1
    Solved = False
End If
If txtG9.Text = "" Or txtG9.Text <> "9" Then
    CTR = CTR + 1
    Solved = False
End If
If txtH1.Text = "" Or txtH1.Text <> "5" Then
    CTR = CTR + 1
    Solved = False
End If
If txtH2.Text = "" Or txtH2.Text <> "4" Then
    CTR = CTR + 1
    Solved = False
End If
If txtH3.Text = "" Or txtH3.Text <> "1" Then
    CTR = CTR + 1
    Solved = False
End If
If txtH4.Text = "" Or txtH4.Text <> "8" Then
    CTR = CTR + 1
    Solved = False
End If
If txtH5.Text = "" Or txtH5.Text <> "9" Then
    CTR = CTR + 1
    Solved = False
End If
If txtI3.Text = "" Or txtI3.Text <> "3" Then
    CTR = CTR + 1
    Solved = False
End If
If txtI4.Text = "" Or txtI4.Text <> "4" Then
    CTR = CTR + 1
    Solved = False
End If
If txtI5.Text = "" Or txtI5.Text <> "4" Then
    CTR = CTR + 1
    Solved = False
End If
If txtI8.Text = "" Or txtI8.Text <> "7" Then
    CTR = CTR + 1
    Solved = False
End If
If txtI9.Text = "" Or txtI9.Text <> "8" Then
    CTR = CTR + 1
    Solved = False
End If
If Solved = True Then
    MsgBox ("Congradulations, you solved the puzzle!")
Else
    If CTR = 1 Then
        MsgBox ("There are " & CTR & " mistake in your puzzle")
    Else
        MsgBox ("There are " & CTR & " mistakes in your puzzle")
    End If
End If
End Sub

Private Sub cmdQuit_Click()
MsgBox ("Good luck with your " & Homework & " hours of homework!")
End
End Sub
