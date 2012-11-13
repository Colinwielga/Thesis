VERSION 5.00
Begin VB.Form frmSecond 
   Caption         =   "Pick Your Board"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8655
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   8595
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Next-->"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4680
         Width           =   1695
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFFC0&
         Caption         =   "<--Back"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4680
         Width           =   1695
      End
      Begin VB.PictureBox picResults1 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "JazzTextExtended"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   6240
         ScaleHeight     =   2235
         ScaleWidth      =   3555
         TabIndex        =   11
         Top             =   1800
         Width           =   3615
      End
      Begin VB.CommandButton cmdPurchaseCapita 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Purchase"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7200
         Width           =   1455
      End
      Begin VB.CommandButton cmdPurchaseRome 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Purchase"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7200
         Width           =   1455
      End
      Begin VB.CommandButton cmdPurchaseForum 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Purchase"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7200
         Width           =   1455
      End
      Begin VB.PictureBox Picture3 
         Height          =   4575
         Left            =   3720
         Picture         =   "Form2.frx":E7B3
         ScaleHeight     =   4515
         ScaleWidth      =   1395
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin VB.PictureBox Picture2 
         Height          =   4575
         Left            =   2040
         Picture         =   "Form2.frx":10BC4
         ScaleHeight     =   4515
         ScaleWidth      =   1395
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
      Begin VB.PictureBox picBoard1 
         Height          =   4575
         Left            =   360
         Picture         =   "Form2.frx":11F28
         ScaleHeight     =   4515
         ScaleWidth      =   1395
         TabIndex        =   2
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "*Must Purchase Item Before Clicking Back To Ensure Correct Price!"
         Height          =   375
         Left            =   5640
         TabIndex        =   17
         Top             =   4200
         Width           =   5175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "$389.99"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "$419.99"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "$319.99"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Forum Insert"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Capita Survival"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Rome Machine"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pick Your Board"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   3000
         TabIndex        =   1
         Top             =   600
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmSecond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdPurchaseForum_Click()
Board = 319.99
If (Money - Board) < 0 Then
    MsgBox "Insufficient Funds... Please Go Back To Start!"
        picResults1.Cls
        frmSecond.Hide
        frmFirst.Show
        cmdPurchaseForum.Enabled = True
        cmdPurchaseCapita.Enabled = True
        cmdPurchaseRome.Enabled = True
Else
        cmdPurchaseForum.Enabled = False
        cmdPurchaseCapita.Enabled = False
        cmdPurchaseRome.Enabled = False
End If
picResults1.Print " Items"; Tab(30); "Expenses"
picResults1.Print "*************************************************"
picResults1.Print "Starting Balance"; Tab(30); FormatCurrency(Money)
picResults1.Print "Forum Insert"; Tab(30); FormatCurrency(-Board)
picResults1.Print "*************************************************"
picResults1.Print "Total"; Tab(30); FormatCurrency(Money - Board)

Money = Money - Board

cmdNext.Enabled = True
cmdBack.Enabled = True
End Sub
Private Sub cmdPurchaseCapita_Click()
Board = 419.99
If (Money - Board) < 0 Then
    MsgBox "Insufficient Funds... Please Go Back To Start!"
        picResults1.Cls
        frmSecond.Hide
        frmFirst.Show
        cmdPurchaseForum.Enabled = True
        cmdPurchaseCapita.Enabled = True
        cmdPurchaseRome.Enabled = True
Else
        cmdPurchaseForum.Enabled = False
        cmdPurchaseCapita.Enabled = False
        cmdPurchaseRome.Enabled = False
End If
picResults1.Print " Items"; Tab(30); "Expenses"
picResults1.Print "*************************************************"
picResults1.Print "Starting Balance"; Tab(30); FormatCurrency(Money)
picResults1.Print "Capita Survival"; Tab(30); FormatCurrency(-Board)
picResults1.Print "*************************************************"
picResults1.Print "Total"; Tab(30); FormatCurrency(Money - Board)

Money = Money - Board

cmdNext.Enabled = True
cmdBack.Enabled = True
End Sub

Private Sub cmdPurchaseRome_Click()
Board = 389.99
If (Money - Board) < 0 Then
    MsgBox "Insufficient Funds... Please Go Back To Start!"
        picResults1.Cls
        frmSecond.Hide
        frmFirst.Show
        cmdPurchaseForum.Enabled = True
        cmdPurchaseCapita.Enabled = True
        cmdPurchaseRome.Enabled = True
Else
        cmdPurchaseForum.Enabled = False
        cmdPurchaseCapita.Enabled = False
        cmdPurchaseRome.Enabled = False
End If
picResults1.Print " Items"; Tab(30); "Expenses"
picResults1.Print "*************************************************"
picResults1.Print "Starting Balance"; Tab(30); FormatCurrency(Money)
picResults1.Print "Rome Machine"; Tab(30); FormatCurrency(-Board)
picResults1.Print "*************************************************"
picResults1.Print "Total"; Tab(30); FormatCurrency(Money - Board)

Money = Money - Board

cmdNext.Enabled = True
cmdBack.Enabled = True
End Sub

Private Sub cmdBack_Click()
picResults1.Cls
Money = Money + Board
cmdPurchaseForum.Enabled = True
cmdPurchaseCapita.Enabled = True
cmdPurchaseRome.Enabled = True

frmSecond.Hide
frmFirst.Show

End Sub

Private Sub cmdNext_Click()
picResults1.Cls

cmdPurchaseForum.Enabled = True
cmdPurchaseCapita.Enabled = True
cmdPurchaseRome.Enabled = True
cmdBack.Enabled = False
frmSecond.Hide
frmThird.Show

End Sub

Private Sub Form_Load()
cmdPurchaseForum.Enabled = True
cmdPurchaseCapita.Enabled = True
cmdPurchaseRome.Enabled = True
End Sub

