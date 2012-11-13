VERSION 5.00
Begin VB.Form frmGWB 
   Caption         =   "Inaugural Speech Mad Lib - George Bush, 2001"
   ClientHeight    =   7500
   ClientLeft      =   2295
   ClientTop       =   2115
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10875
   Begin VB.PictureBox picAmericanGWB 
      Height          =   7695
      Left            =   -360
      Picture         =   "frmGWMadLibs.frx":0000
      ScaleHeight     =   7635
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   -120
      Width           =   11655
      Begin VB.CommandButton cmdGoBack 
         Caption         =   "Go Back"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   24
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdRead 
         Caption         =   "Input Words"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   23
         Top             =   6120
         Width           =   2055
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "Display Words"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         TabIndex        =   22
         Top             =   6120
         Width           =   2055
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   21
         Top             =   6960
         Width           =   1095
      End
      Begin VB.PictureBox pic11 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   19
         Top             =   4440
         Width           =   855
      End
      Begin VB.PictureBox pic12 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   18
         Top             =   4440
         Width           =   855
      End
      Begin VB.PictureBox pic13 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   17
         Top             =   5040
         Width           =   975
      End
      Begin VB.PictureBox pic14 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   16
         Top             =   5280
         Width           =   855
      End
      Begin VB.PictureBox pic6 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   14
         Top             =   3120
         Width           =   855
      End
      Begin VB.PictureBox pic7 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9720
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   13
         Top             =   3720
         Width           =   975
      End
      Begin VB.PictureBox pic8 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   12
         Top             =   3960
         Width           =   855
      End
      Begin VB.PictureBox pic9 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   11
         Top             =   3960
         Width           =   855
      End
      Begin VB.PictureBox pic10 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   10
         Top             =   3960
         Width           =   855
      End
      Begin VB.PictureBox pic4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.PictureBox pic5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         ScaleHeight     =   195
         ScaleWidth      =   1035
         TabIndex        =   6
         Top             =   2400
         Width           =   1095
      End
      Begin VB.PictureBox pic2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox pic3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         ScaleHeight     =   195
         ScaleWidth      =   795
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.PictureBox pic1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lbl6 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmGWMadLibs.frx":63E15
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   840
         TabIndex        =   20
         Top             =   4440
         Visible         =   0   'False
         Width           =   9375
      End
      Begin VB.Label lbl5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmGWMadLibs.frx":63FEC
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   840
         TabIndex        =   15
         Top             =   3120
         Visible         =   0   'False
         Width           =   9735
      End
      Begin VB.Label lbl3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmGWMadLibs.frx":64194
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.Label lbl4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmGWMadLibs.frx":64243
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   8
         Top             =   2400
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.Label lbl2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmGWMadLibs.frx":642CD
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.Label lbl1 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "President Clinton, distinguished guests and my fellow                   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   7575
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmGWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As String, a1 As String, a2 As String, a3 As String, a4 As String
Dim a5 As String, a6 As String, a7 As String, a8 As String, a9 As String, a10 As String
Dim a11 As String, a12 As String, a13 As String, a14 As String
Private Sub cmdGoBack_Click()
    frmBeginMadLib.Show
    frmGWB.Hide
End Sub

Private Sub cmdRead_Click()
    cmdDisplay.Enabled = True
    a1 = InputBox("Enter an animal. (plural)")
    a2 = InputBox("Enter an adjective.")
    a3 = InputBox("Enter a verb.")
    a4 = InputBox("Enter a snack food.")
    a5 = InputBox("Enter a profession.")
    a6 = InputBox("Enter a verb.")
    a7 = InputBox("Enter two sets of opposing verbs.  (i.e. walk and run, or sleep and exercise)")
    a8 = InputBox("Enter the opposing verb to the last word.")
    a9 = InputBox("Enter the last set of opposing verbs.")
    a10 = InputBox("Enter the opposing verb to the previous word.")
    a11 = InputBox("Enter an adjective.")
    a12 = InputBox("Enter another adjective.")
    a13 = InputBox("Enter an adjective describing a person.")
    a14 = InputBox("Enter a verb.")
    MsgBox ("Good Job, when you are ready to see your masterpiece, CLICK on Display Words.")


End Sub


Private Sub cmdDisplay_Click()
    lbl1.Visible = True
    lbl2.Visible = True
    lbl3.Visible = True
    lbl4.Visible = True
    lbl5.Visible = True
    lbl6.Visible = True
    pic1.Print a1
    pic2.Print a2
    pic3.Print a3
    pic4.Print a4
    pic5.Print a5
    pic6.Print a6
    pic7.Print a7
    pic8.Print a8
    pic9.Print a9
    pic10.Print a10
    pic11.Print a11
    pic12.Print a12
    pic13.Print a13
    pic14.Print a14
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

