VERSION 5.00
Begin VB.Form DiveSheet 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15195
   FillColor       =   &H8000000A&
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   15195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOpposingSchool 
      Height          =   375
      Left            =   5640
      TabIndex        =   140
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox txtSchool 
      Height          =   375
      Left            =   1200
      TabIndex        =   139
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox txtDiverName 
      Height          =   375
      Left            =   1200
      TabIndex        =   137
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox txtJudge5_11 
      Height          =   495
      Left            =   10800
      TabIndex        =   134
      Top             =   7800
      Width           =   615
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      Height          =   495
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtJudge5_10 
      Height          =   495
      Left            =   10800
      TabIndex        =   132
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtJudge5_9 
      Height          =   495
      Left            =   10800
      TabIndex        =   131
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox txtJudge5_8 
      Height          =   495
      Left            =   10800
      TabIndex        =   130
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtJudge5_7 
      Height          =   495
      Left            =   10800
      TabIndex        =   129
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtJudge5_6 
      Height          =   495
      Left            =   10800
      TabIndex        =   128
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtJudge5_5 
      Height          =   495
      Left            =   10800
      TabIndex        =   127
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtJudge5_4 
      Height          =   495
      Left            =   10800
      TabIndex        =   126
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtJudge5_3 
      Height          =   495
      Left            =   10800
      TabIndex        =   125
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_11 
      Height          =   495
      Left            =   9960
      TabIndex        =   124
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_10 
      Height          =   495
      Left            =   9960
      TabIndex        =   123
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_9 
      Height          =   495
      Left            =   9960
      TabIndex        =   122
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_8 
      Height          =   495
      Left            =   9960
      TabIndex        =   121
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_7 
      Height          =   495
      Left            =   9960
      TabIndex        =   120
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_6 
      Height          =   495
      Left            =   9960
      TabIndex        =   119
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_5 
      Height          =   495
      Left            =   9960
      TabIndex        =   118
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_4 
      Height          =   495
      Left            =   9960
      TabIndex        =   117
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_3 
      Height          =   495
      Left            =   9960
      TabIndex        =   116
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdTotalScore 
      BackColor       =   &H00C000C0&
      Caption         =   "Total Score"
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdTotalDD 
      BackColor       =   &H00C000C0&
      Caption         =   "Total Required DD"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   8520
      Width           =   1695
   End
   Begin VB.TextBox txtJudge3_11 
      Height          =   495
      Left            =   9120
      TabIndex        =   113
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox txtJudge3_10 
      Height          =   495
      Left            =   9120
      TabIndex        =   112
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtJudge3_9 
      Height          =   495
      Left            =   9120
      TabIndex        =   111
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox txtJudge3_8 
      Height          =   495
      Left            =   9120
      TabIndex        =   110
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtJudge3_7 
      Height          =   495
      Left            =   9120
      TabIndex        =   109
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtJudge3_6 
      Height          =   495
      Left            =   9120
      TabIndex        =   108
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtJudge3_5 
      Height          =   495
      Left            =   9120
      TabIndex        =   107
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtJudge3_4 
      Height          =   495
      Left            =   9120
      TabIndex        =   106
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtJudge3_3 
      Height          =   495
      Left            =   9120
      TabIndex        =   105
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_11 
      Height          =   495
      Left            =   8280
      TabIndex        =   104
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_10 
      Height          =   495
      Left            =   8280
      TabIndex        =   103
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_9 
      Height          =   495
      Left            =   8280
      TabIndex        =   102
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_8 
      Height          =   495
      Left            =   8280
      TabIndex        =   101
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_7 
      Height          =   495
      Left            =   8280
      TabIndex        =   100
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_6 
      Height          =   495
      Left            =   8280
      TabIndex        =   99
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_5 
      Height          =   495
      Left            =   8280
      TabIndex        =   98
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_4 
      Height          =   495
      Left            =   8280
      TabIndex        =   97
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_3 
      Height          =   495
      Left            =   8280
      TabIndex        =   96
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_11 
      Height          =   495
      Left            =   7440
      TabIndex        =   95
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_10 
      Height          =   495
      Left            =   7440
      TabIndex        =   94
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_9 
      Height          =   495
      Left            =   7440
      TabIndex        =   93
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_8 
      Height          =   495
      Left            =   7440
      TabIndex        =   92
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_7 
      Height          =   495
      Left            =   7440
      TabIndex        =   91
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_6 
      Height          =   495
      Left            =   7440
      TabIndex        =   90
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_5 
      Height          =   495
      Left            =   7440
      TabIndex        =   89
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_4 
      Height          =   495
      Left            =   7440
      TabIndex        =   88
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_3 
      Height          =   495
      Left            =   7440
      TabIndex        =   87
      Top             =   3000
      Width           =   615
   End
   Begin VB.PictureBox picReqDD 
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   86
      Top             =   8520
      Width           =   615
   End
   Begin VB.PictureBox picDes11 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   85
      Top             =   7800
      Width           =   3375
   End
   Begin VB.PictureBox picDes10 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   84
      Top             =   7200
      Width           =   3375
   End
   Begin VB.PictureBox picDes9 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   83
      Top             =   6600
      Width           =   3375
   End
   Begin VB.PictureBox picDes8 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   82
      Top             =   6000
      Width           =   3375
   End
   Begin VB.PictureBox picDes7 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   81
      Top             =   5400
      Width           =   3375
   End
   Begin VB.PictureBox picDes6 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   80
      Top             =   4800
      Width           =   3375
   End
   Begin VB.PictureBox picDes5 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   79
      Top             =   4200
      Width           =   3375
   End
   Begin VB.PictureBox picDes4 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   78
      Top             =   3600
      Width           =   3375
   End
   Begin VB.PictureBox picDes3 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   77
      Top             =   3000
      Width           =   3375
   End
   Begin VB.PictureBox picDes2 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   76
      Top             =   2400
      Width           =   3375
   End
   Begin VB.PictureBox picDes1 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   3315
      TabIndex        =   75
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtJudge5_2 
      Height          =   495
      Left            =   10800
      TabIndex        =   74
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_2 
      Height          =   495
      Left            =   9960
      TabIndex        =   73
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtJudge3_2 
      Height          =   495
      Left            =   9120
      TabIndex        =   72
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_2 
      Height          =   495
      Left            =   8280
      TabIndex        =   71
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_2 
      Height          =   495
      Left            =   7440
      TabIndex        =   70
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtJudge5_1 
      Height          =   495
      Left            =   10800
      TabIndex        =   67
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtJudge4_1 
      Height          =   495
      Left            =   9960
      TabIndex        =   66
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text22 
      Height          =   495
      Left            =   5880
      TabIndex        =   64
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox Text21 
      Height          =   495
      Left            =   5880
      TabIndex        =   63
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox Text20 
      Height          =   495
      Left            =   5880
      TabIndex        =   62
      Top             =   7800
      Width           =   495
   End
   Begin VB.TextBox Text19 
      Height          =   495
      Left            =   5880
      TabIndex        =   61
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox Text18 
      Height          =   495
      Left            =   5880
      TabIndex        =   60
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   5880
      TabIndex        =   59
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox Text16 
      Height          =   495
      Left            =   5880
      TabIndex        =   58
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   5880
      TabIndex        =   57
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   5880
      TabIndex        =   56
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   5880
      TabIndex        =   55
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   5880
      TabIndex        =   54
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtDD11 
      Height          =   495
      Left            =   6600
      TabIndex        =   53
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox txtDD10 
      Height          =   495
      Left            =   6600
      TabIndex        =   52
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtDD9 
      Height          =   495
      Left            =   6600
      TabIndex        =   51
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox txtDD8 
      Height          =   495
      Left            =   6600
      TabIndex        =   50
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtDD7 
      Height          =   495
      Left            =   6600
      TabIndex        =   49
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtDD6 
      Height          =   495
      Left            =   6600
      TabIndex        =   48
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtDD5 
      Height          =   495
      Left            =   6600
      TabIndex        =   47
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtDD4 
      Height          =   495
      Left            =   6600
      TabIndex        =   46
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtDD3 
      Height          =   495
      Left            =   6600
      TabIndex        =   45
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtDD2 
      Height          =   495
      Left            =   6600
      TabIndex        =   43
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtDD1 
      Height          =   495
      Left            =   6600
      TabIndex        =   42
      Top             =   1800
      Width           =   615
   End
   Begin VB.PictureBox pic_Score11 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   41
      Top             =   7800
      Width           =   1335
   End
   Begin VB.PictureBox pic_Score10 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   40
      Top             =   7200
      Width           =   1335
   End
   Begin VB.PictureBox pic_Score9 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   39
      Top             =   6600
      Width           =   1335
   End
   Begin VB.PictureBox pic_Score8 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   38
      Top             =   6000
      Width           =   1335
   End
   Begin VB.PictureBox pic_Score7 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   37
      Top             =   5400
      Width           =   1335
   End
   Begin VB.PictureBox pic_Score6 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   36
      Top             =   4800
      Width           =   1335
   End
   Begin VB.PictureBox pic_Score5 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   35
      Top             =   4200
      Width           =   1335
   End
   Begin VB.PictureBox pic_Score4 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   34
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox pic_Score3 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   33
      Top             =   3000
      Width           =   1335
   End
   Begin VB.PictureBox pic_Score2 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   32
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton ComputeDive11 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 11 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton ComputeDive10 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 10 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton ComputeDive9 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 9 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton ComputeDive8 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 8 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton ComputeDive7 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 7 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton ComputeDive6 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 6 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton ComputeDive5 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 5 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton ComputeDive4 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 4 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton ComputeDive3 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 3 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton ComputeDive2 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 2 Score"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox pic_total 
      Height          =   615
      Left            =   13800
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   21
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox txtDive11 
      Height          =   495
      Left            =   1200
      TabIndex        =   20
      Top             =   7800
      Width           =   855
   End
   Begin VB.TextBox txtDive10 
      Height          =   495
      Left            =   1200
      TabIndex        =   19
      Top             =   7200
      Width           =   855
   End
   Begin VB.TextBox txtDive9 
      Height          =   495
      Left            =   1200
      TabIndex        =   18
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox txtDive8 
      Height          =   495
      Left            =   1200
      TabIndex        =   17
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtDive7 
      Height          =   495
      Left            =   1200
      TabIndex        =   16
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtDive6 
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtDive5 
      Height          =   495
      Left            =   1200
      TabIndex        =   14
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtDive4 
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtDive3 
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtDive2 
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.PictureBox pic_Score1 
      Height          =   495
      Left            =   13800
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton ComputeDive1 
      BackColor       =   &H0000FF00&
      Caption         =   "Compute Dive 1 Scores"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtDive1 
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtJudge3_1 
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtJudge2_1 
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtJudge1_1 
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      Caption         =   "VS."
      Height          =   255
      Left            =   5160
      TabIndex        =   142
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "SCHOOL"
      Height          =   255
      Left            =   0
      TabIndex        =   141
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      Caption         =   "DIVER NAME"
      Height          =   255
      Left            =   0
      TabIndex        =   138
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "REQUIREDS(5)"
      Height          =   255
      Left            =   0
      TabIndex        =   136
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "OPTIONALS(6)"
      Height          =   255
      Left            =   0
      TabIndex        =   135
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Judge 5"
      Height          =   255
      Index           =   4
      Left            =   10800
      TabIndex        =   69
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Judge 4"
      Height          =   255
      Index           =   3
      Left            =   9960
      TabIndex        =   68
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "       Pos.          (t, p, s, f)"
      Height          =   495
      Left            =   5640
      TabIndex        =   65
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label DegreeOfDifficulty 
      BackColor       =   &H0080FFFF&
      Caption         =   "D.D."
      Height          =   255
      Left            =   6720
      TabIndex        =   44
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label DiveDescription 
      BackColor       =   &H00FFFF00&
      Caption         =   "Dive Desctiption"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label DiveNo 
      BackColor       =   &H0080FFFF&
      Caption         =   "Dive #"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Judge 3"
      Height          =   255
      Index           =   2
      Left            =   9120
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Judge 2"
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Judge 1"
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "DiveSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Competitive Diving Form
'Form Name: DiveSheet
'Marcus Rien
'3/22/09
'This form gathers data from the user (i.e. dive information and judges scores) and computes the each dive's score
'This form also computes the Required Dive's Degree of Difficulty as well as teh Total Score of all of the Dives

Dim I As Integer
Dim CTR As Integer
Dim found As Boolean
Dim DiveNumber(1 To 200) As String, DiveDD(1 To 200) As Single, DiveDescript(1 To 200) As String

Dim DD1 As Single, DD2 As Single, DD3 As Single, DD4 As Single, DD5 As Single
Dim DD6 As Single, DD7 As Single, DD8 As Single, DD9 As Single, DD10 As Single, DD11 As Single

Dim Dive1 As String, Dive2 As String, Dive3 As String, Dive4 As String
Dim Dive5 As String, Dive6 As String, Dive7 As String, Dive8 As String
Dim Dive9 As String, Dive10 As String, Dive11 As String
Dim SumScore(1 To 11) As Single
'This adds up the total of all of the dives
Private Sub cmdTotalScore_Click()
pic_total.Cls
Dim TotalScore As Double
For I = 1 To 11
TotalScore = TotalScore + SumScore(I)
Next I
pic_total.Print FormatNumber(TotalScore, 2)
If TotalScore > 435 Then
    frmNationalCut.Show
    DiveSheet.Hide
End If

End Sub
'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 2nd Dive another form opens to help them.
Private Sub ComputeDive1_Click()
pic_Score1.Cls


Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive1 = txtDive1.Text
DD1 = txtDD1.Text
picDes1.Cls
Close #1
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR

    If Dive1 = DiveNumber(I) Then
        If DD1 <> DiveDD(I) Then
        'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        txtDD1.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive1 & "." & " Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD1.Text = "HELP ME" Then
                txtDD1.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes1.Print DiveDescript(I)
        Else
        picDes1.Print DiveDescript(I)
        
        End If
    End If
Next
If txtJudge1_1.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore1 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_1.Text
score2 = txtJudge2_1.Text
score3 = txtJudge3_1.Text
score4 = txtJudge4_1.Text
score5 = txtJudge5_1.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore1 = sum * txtDD1.Text
SumScore(1) = totalScore1
pic_Score1.Print totalScore1
End If
End Sub

'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 2nd Dive another form opens to help them.
Private Sub ComputeDive2_Click()
pic_Score2.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty
Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive2 = txtDive2.Text
DD2 = txtDD2.Text
picDes2.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive2 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD2 <> DiveDD(I) Then
        txtDD2.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive2 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD2.Text = "HELP ME" Then
                txtDD2.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes2.Print DiveDescript(I)
        Else
        picDes2.Print DiveDescript(I)
        End If
    End If
Next

If txtJudge1_2.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore2 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_2.Text
score2 = txtJudge2_2.Text
score3 = txtJudge3_2.Text
score4 = txtJudge4_2.Text
score5 = txtJudge5_2.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore2 = sum * txtDD2.Text
SumScore(2) = totalScore2
pic_Score2.Print totalScore2
End If
Close #1
End Sub

'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 3rd Dive another form opens to help them.
Private Sub ComputeDive3_Click()
pic_Score3.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive3 = txtDive3.Text
DD3 = txtDD3.Text
picDes3.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive3 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD3 <> DiveDD(I) Then
        txtDD3.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive3 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD3.Text = "HELP ME" Then
                txtDD3.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes3.Print DiveDescript(I)
        Else
        picDes3.Print DiveDescript(I)
        End If
    End If
Next

If txtJudge1_3.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore3 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_3.Text
score2 = txtJudge2_3.Text
score3 = txtJudge3_3.Text
score4 = txtJudge4_3.Text
score5 = txtJudge5_3.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore3 = sum * txtDD3.Text
SumScore(3) = totalScore3
pic_Score3.Print totalScore3
End If

Close #1
End Sub
'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 4th Dive another form opens to help them.
Private Sub ComputeDive4_Click()
pic_Score4.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive4 = txtDive4.Text
DD4 = txtDD4.Text
picDes4.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive4 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD4 <> DiveDD(I) Then
        txtDD4.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive4 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD4.Text = "HELP ME" Then
                txtDD4.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes4.Print DiveDescript(I)
        Else
        picDes4.Print DiveDescript(I)
        End If
    End If
Next
If txtJudge1_4.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore4 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_4.Text
score2 = txtJudge2_4.Text
score3 = txtJudge3_4.Text
score4 = txtJudge4_4.Text
score5 = txtJudge5_4.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore4 = sum * txtDD4.Text
SumScore(4) = totalScore4
pic_Score4.Print totalScore4
End If

Close #1
End Sub

'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 5th Dive another form opens to help them.
Private Sub ComputeDive5_Click()
pic_Score5.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive5 = txtDive5.Text
DD5 = txtDD5.Text
picDes5.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive5 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD5 <> DiveDD(I) Then
        txtDD5.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive5 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD5.Text = "HELP ME" Then
                txtDD5.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes5.Print DiveDescript(I)
        Else
        picDes5.Print DiveDescript(I)
        End If
    End If
Next

If txtJudge1_5.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore5 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_5.Text
score2 = txtJudge2_5.Text
score3 = txtJudge3_5.Text
score4 = txtJudge4_5.Text
score5 = txtJudge5_5.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore5 = sum * txtDD5.Text
SumScore(5) = totalScore5
pic_Score5.Print totalScore5
End If

Close #1
End Sub

'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 6th Dive another form opens to help them.
Private Sub ComputeDive6_Click()
pic_Score6.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive6 = txtDive6.Text
DD6 = txtDD6.Text
picDes6.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive6 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD6 <> DiveDD(I) Then
        txtDD6.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive6 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD6.Text = "HELP ME" Then
                txtDD6.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes6.Print DiveDescript(I)
        Else
        picDes6.Print DiveDescript(I)
        End If
    End If
Next

If txtJudge1_6.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore6 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_6.Text
score2 = txtJudge2_6.Text
score3 = txtJudge3_6.Text
score4 = txtJudge4_6.Text
score5 = txtJudge5_6.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore6 = sum * txtDD6.Text
SumScore(6) = totalScore6
pic_Score6.Print totalScore6

End If
Close #1
End Sub

'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 7th Dive another form opens to help them.
Private Sub ComputeDive7_Click()
pic_Score7.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive7 = txtDive7.Text
DD7 = txtDD7.Text
picDes7.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive7 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD7 <> DiveDD(I) Then
        txtDD7.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive7 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD7.Text = "HELP ME" Then
                txtDD7.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes7.Print DiveDescript(I)
        Else
        picDes7.Print DiveDescript(I)
        End If
    End If
Next
Close #1

If txtJudge1_7.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore7 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_7.Text
score2 = txtJudge2_7.Text
score3 = txtJudge3_7.Text
score4 = txtJudge4_7.Text
score5 = txtJudge5_7.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore7 = sum * txtDD7.Text
SumScore(7) = totalScore7
pic_Score7.Print totalScore7
End If

End Sub

'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 8th Dive another form opens to help them.
Private Sub ComputeDive8_Click()
pic_Score8.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive8 = txtDive8.Text
DD8 = txtDD8.Text
picDes8.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive8 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD8 <> DiveDD(I) Then
        txtDD8.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive8 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD8.Text = "HELP ME" Then
                txtDD9.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes8.Print DiveDescript(I)
        Else
        picDes8.Print DiveDescript(I)
        End If
    End If
Next

If txtJudge1_8.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore8 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_8.Text
score2 = txtJudge2_8.Text
score3 = txtJudge3_8.Text
score4 = txtJudge4_8.Text
score5 = txtJudge5_8.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore8 = sum * txtDD8.Text
SumScore(8) = totalScore8
pic_Score8.Print totalScore8
End If

Close #1
End Sub

'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 9th Dive another form opens to help them.
Private Sub ComputeDive9_Click()
pic_Score9.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive9 = txtDive9.Text
DD9 = txtDD9.Text
picDes9.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive9 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD9 <> DiveDD(I) Then
        txtDD9.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive9 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD9.Text = "HELP ME" Then
                txtDD9.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes9.Print DiveDescript(I)
        Else
        picDes9.Print DiveDescript(I)
        End If
    End If
Next

If txtJudge1_9.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore9 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_9.Text
score2 = txtJudge2_9.Text
score3 = txtJudge3_9.Text
score4 = txtJudge4_9.Text
score5 = txtJudge5_9.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore9 = sum * txtDD9.Text
SumScore(9) = totalScore9
pic_Score9.Print totalScore9
End If

Close #1
End Sub

'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 10th Dive another form opens to help them.
Private Sub ComputeDive10_Click()
pic_Score10.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive10 = txtDive10.Text
DD10 = txtDD10.Text
picDes10.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive10 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD10 <> DiveDD(I) Then
        txtDD10.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive10 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD10.Text = "HELP ME" Then
                txtDD11.Text = " "
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes10.Print DiveDescript(I)
        Else
        picDes10.Print DiveDescript(I)
        End If
    End If
Next

If txtJudge1_10.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore10 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_10.Text
score2 = txtJudge2_10.Text
score3 = txtJudge3_10.Text
score4 = txtJudge4_10.Text
score5 = txtJudge5_10.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore10 = sum * txtDD10.Text
SumScore(10) = totalScore10
pic_Score10.Print totalScore10
End If

Close #1
End Sub

'This matches the Dive Number with the Dive Degree of difficulty and prints the description.
'If the Dive needs help with their Degree of Difficulty input for the 11th Dive another form opens to help them.
Private Sub ComputeDive11_Click()
pic_Score11.Cls

Open App.Path & "\1meterDD.txt" For Input As #1
CTR = 0
'Dives and Degree of Difficulty

Do Until EOF(1)
    CTR = CTR + 1
       Input #1, DiveNumber(CTR), DiveDD(CTR), DiveDescript(CTR)
Loop
Dive11 = txtDive11.Text
DD11 = txtDD11.Text
picDes11.Cls
'This makes sure that the Degree of difficulty is equal to the dive number
For I = 1 To CTR
    If Dive11 = DiveNumber(I) Then
    'This lets the user know that they have entered the wrong Degree of Difficulty for that Dive
        If DD11 <> DiveDD(I) Then
        txtDD11.Text = InputBox("Are you sure the Degree of Difficulty is correct? Please re-enter the Degree of Difficulty for " & Dive11 & "." & "                      Note: For help on the correct Degree of Difficulty, enter HELP ME" & ".")
            If txtDD11.Text = "HELP ME" Then
                DiveSheet.Hide
                DiveList.Show
            End If
        picDes11.Print DiveDescript(I)
        Else
        picDes11.Print DiveDescript(I)
        End If
    End If
Next

If txtJudge1_11.Text <> "" Then

'JUDGES SCORES
Dim Scores(1 To 5) As Single
Dim largest As Single
Dim smallest As Single
Dim sum As Single
Dim score1 As Single
Dim score2 As Single
Dim score3 As Single
Dim score4 As Single
Dim score5 As Single
Dim totalScore11 As Single
Dim found1 As Integer
Dim found2 As Integer

score1 = txtJudge1_11.Text
score2 = txtJudge2_11.Text
score3 = txtJudge3_11.Text
score4 = txtJudge4_11.Text
score5 = txtJudge5_11.Text

Scores(1) = score1
Scores(2) = score2
Scores(3) = score3
Scores(4) = score4
Scores(5) = score5

sum = 0
largest = 0
smallest = 10
'This sorts through the judges scores and eliminates the high and low number score
For I = 1 To 5
If Scores(I) > largest Then
    largest = Scores(I)
    found1 = I
    ElseIf Scores(I) < smallest Then
    found2 = I
    smallest = Scores(I)
End If
Next I
'This sums up the remaining judge scores
For I = 1 To 5
If I <> found1 Then
    If I <> found2 Then
    sum = sum + Scores(I)
    End If
End If
Next I

totalScore11 = sum * txtDD11.Text
SumScore(11) = totalScore11
pic_Score11.Print totalScore11
End If

Close #1
End Sub

'This command adds up the Total Degree of Difficulty (DD) for the required dives
'If the DD is more than 9.0 it lets the user know that he/she needs to alter their required dives
Private Sub cmdTotalDD_Click()
picReqDD.Cls
Dim totalReqDD As Single
totalReqDD = txtDD7.Text
totalReqDD = totalReqDD + txtDD8.Text
totalReqDD = totalReqDD + txtDD9.Text
totalReqDD = totalReqDD + txtDD10.Text
totalReqDD = totalReqDD + txtDD11.Text
If totalReqDD > 9# Then
MsgBox "Your Required Dives may not be in excess of 9.0 Degree of Difficulty. Please alter your REQUIREDS.", , "ERROR!"
Else
picReqDD.Print totalReqDD
End If
End Sub
