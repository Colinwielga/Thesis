VERSION 5.00
Begin VB.Form frmpictures 
   BackColor       =   &H00000000&
   Caption         =   "Roster Pictures"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Navigate"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6840
      Width           =   1455
   End
   Begin VB.PictureBox Picture13 
      Height          =   2295
      Left            =   4200
      Picture         =   "BUFFS.vpw.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   12
      Top             =   5880
      Width           =   1815
   End
   Begin VB.PictureBox Picture12 
      Height          =   2295
      Left            =   6240
      Picture         =   "BUFFS.vpw.frx":2624
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   11
      Top             =   5880
      Width           =   1815
   End
   Begin VB.PictureBox Picture11 
      Height          =   2295
      Left            =   2160
      Picture         =   "BUFFS.vpw.frx":4B4E
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   10
      Top             =   5880
      Width           =   1815
   End
   Begin VB.PictureBox Picture10 
      Height          =   2295
      Left            =   8280
      Picture         =   "BUFFS.vpw.frx":5F44
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox Picture9 
      Height          =   2295
      Left            =   6240
      Picture         =   "BUFFS.vpw.frx":8C6E
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox Picture8 
      Height          =   2295
      Left            =   4200
      Picture         =   "BUFFS.vpw.frx":D062
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   7
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox Picture7 
      Height          =   2295
      Left            =   120
      Picture         =   "BUFFS.vpw.frx":E311
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox Picture6 
      Height          =   2295
      Left            =   2160
      Picture         =   "BUFFS.vpw.frx":12ADD
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox Picture4 
      Height          =   2295
      Left            =   8280
      Picture         =   "BUFFS.vpw.frx":17248
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture5 
      Height          =   2295
      Left            =   6240
      Picture         =   "BUFFS.vpw.frx":185CD
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   4200
      Picture         =   "BUFFS.vpw.frx":1CCA7
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   2160
      Picture         =   "BUFFS.vpw.frx":1F218
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   120
      Picture         =   "BUFFS.vpw.frx":21F14
      ScaleHeight     =   2235
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblWirt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "    Randie Wirt           No. 54"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   25
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label lblWaner 
      BackColor       =   &H00FFFFFF&
      Caption         =   "    Emily Waner           No. 4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   24
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label lblNedovic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   Anna Nedovic          No. 12"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   23
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label lblMetoyer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  Amber Metoyer          No. 2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   22
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblJohns 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Veronica Johns-R.         No. 5"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblJones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   Cecily Jones            No. 55"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   20
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblLaw 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   Whitney Law           No. 24"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   19
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblLini 
      BackColor       =   &H00FFFFFF&
      Caption         =   "    Sarah Lini             No. 34"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblIlic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   Jasmina Ilic            No. 21"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   17
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblBjorklund 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   Tera Bjorklund           No. 50"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblFagan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "    Kate Fagan              No. 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblHoward 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  Leslie Howard          No. 14"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblBillingsley 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Maria Billingsley        No. 10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "frmpictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label10_Click()

End Sub

Private Sub cmdNext_Click()
frmpictures.Hide
frmteamstats.Show
End Sub
