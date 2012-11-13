VERSION 5.00
Begin VB.Form frmTrivia 
   BackColor       =   &H00400000&
   Caption         =   "Twins Trivia"
   ClientHeight    =   8580
   ClientLeft      =   2580
   ClientTop       =   870
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   10515
   Begin VB.CommandButton cmdAnswers 
      BackColor       =   &H000000C0&
      Caption         =   "Show Answers"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   8040
      Width           =   2055
   End
   Begin VB.TextBox txtAns10 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   62
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtAns9 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   57
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtAns8 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   52
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtAns7 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   47
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox txtAns6 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   42
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtAns5 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   37
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtAns4 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   32
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtAns3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   27
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtAns2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   22
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox txtAns1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit Twins Territory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000C0&
      Caption         =   "Back to Twins Territory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdFinish 
      BackColor       =   &H000000C0&
      Caption         =   "Finished!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label lbl10d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) Pirahnas"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   61
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label lbl10b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B) Sharks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   60
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lbl10c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Crappies"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   59
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label lbl10a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) Swordfish"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   58
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label lbl9d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) Dan Gladden"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   56
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label lbl9b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B) Torii Hunter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   55
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lbl9c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Jarvis Brown"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   54
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label lbl9a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) Kirby Puckett"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   53
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lbl8d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) Dwight Lowry"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   51
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lbl8b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B) Kent Hrbek"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   50
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lbl8c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Bert Blyleven"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   49
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lbl8a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) Harmon Killebrew"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   48
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lbl7d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) Torii Hunter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   46
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lbl7b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B) Joe Mauer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   45
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lbl7c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Johan Santana"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   44
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lbl7a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) Justin Morneau"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   43
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lbl6d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) Kent Hrbek"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   41
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lbl6b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B) Torii Hunter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   40
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lbl6c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Kirby Puckett"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   39
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lbl6a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) Harmon Killebrew"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   38
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lbl5d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) Four"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   36
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label lbl5b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Three"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   35
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lbl5c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Three"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label lbl5a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) One"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label lbl4d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) Brad Radke"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   31
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label lbl4b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Frank Viola"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   30
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lbl4c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Bert Blyleven"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label lbl4a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) Johan Santana"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lbl3d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) 1999"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lbl3b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B) 2002"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   25
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lbl3c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) 2000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lbl3a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) 2005"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lbl2d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) Torii Hunter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lbl2b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B) Joe Mauer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lbl2c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) Johan Santana"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lbl2a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) Justin Morneau"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lbl1a 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A) Cleveland Cheifs"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lbl1d 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D) Minneapolis Moondogs"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lbl1b 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B) Washington Senators"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lbl1c 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C) St. Paul Saints"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblQ4 
      BackColor       =   &H000000C0&
      Caption         =   "4) Who leads the Twins franchise in most strike-outs in one season with 265?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   4695
   End
   Begin VB.Label lblQ10 
      BackColor       =   &H000000C0&
      Caption         =   "10) During the 2006 season, the Twins earned the nickname of what type of fish?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   6480
      Width           =   4695
   End
   Begin VB.Label lblQ9 
      BackColor       =   &H000000C0&
      Caption         =   "9) Who was the Twins starting center fielder in the 1991 World Series?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   4920
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "8) Which former Twin now has his own TV show about the great outdoors?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label lblQ7 
      BackColor       =   &H000000C0&
      Caption         =   "7) Which Twin earned the Title of the 2006 American League Batting Champ?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label lblQ6 
      BackColor       =   &H000000C0&
      Caption         =   "6) Who holds the franchise record for most RBIs?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblQ5 
      BackColor       =   &H000000C0&
      Caption         =   "5) How many times did the Twins win the Division Series between 2002 and 2007?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   6480
      Width           =   4695
   End
   Begin VB.Label lblQ3 
      BackColor       =   &H000000C0&
      Caption         =   "3) In what year did Ron Gardenhire become the new manager for the Twins?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label lblQ2 
      BackColor       =   &H000000C0&
      Caption         =   "2)Which Twin was named the 2006 American League MVP?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label lblQ1 
      BackColor       =   &H000000C0&
      Caption         =   "1) What was the original name of the Twins franchise?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim correct As Integer

Private Sub cmdAnswers_Click()
'this command will disable the visibility of all the wrong answers
lbl1a.Visible = False
lbl1c.Visible = False
lbl1d.Visible = False
lbl2b.Visible = False
lbl2c.Visible = False
lbl2d.Visible = False
lbl3a.Visible = False
lbl3c.Visible = False
lbl3d.Visible = False
lbl4b.Visible = False
lbl4c.Visible = False
lbl4d.Visible = False
lbl5a.Visible = False
lbl5b.Visible = False
lbl5c.Visible = False
lbl6b.Visible = False
lbl6c.Visible = False
lbl6d.Visible = False
lbl7a.Visible = False
lbl7c.Visible = False
lbl7d.Visible = False
lbl8a.Visible = False
lbl8c.Visible = False
lbl8d.Visible = False
lbl9b.Visible = False
lbl9c.Visible = False
lbl9d.Visible = False
lbl10a.Visible = False
lbl10b.Visible = False
lbl10c.Visible = False

'disables finised button
cmdFinish.Enabled = False

End Sub

Private Sub cmdBack_Click()
    frmTrivia.Hide 'hides trivia form
    frmMain.Show 'goes back to main form
End Sub

Private Sub cmdExit_Click()
    End 'ends program
End Sub


Private Sub cmdFinish_Click()
'initialize correct
correct = 0
'determines the number of answers answered correctly
If txtAns1.Text = "B" Then correct = correct + 1
If txtAns2.Text = "A" Then correct = correct + 1
If txtAns3.Text = "B" Then correct = correct + 1
If txtAns4.Text = "A" Then correct = correct + 1
If txtAns5.Text = "D" Then correct = correct + 1
If txtAns6.Text = "A" Then correct = correct + 1
If txtAns7.Text = "B" Then correct = correct + 1
If txtAns8.Text = "B" Then correct = correct + 1
If txtAns9.Text = "A" Then correct = correct + 1
If txtAns10.Text = "D" Then correct = correct + 1

'display Results
MsgBox "You answered " & correct & " questions correctly", , "Results"

cmdAnswers.Enabled = True 'enables answers command
End Sub

