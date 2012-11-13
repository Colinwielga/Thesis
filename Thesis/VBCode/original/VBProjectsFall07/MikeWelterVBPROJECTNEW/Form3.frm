VERSION 5.00
Begin VB.Form frmThird 
   Caption         =   "Pick Your Bindings"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   11895
      Left            =   -1320
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   11835
      ScaleWidth      =   15795
      TabIndex        =   0
      Top             =   -1320
      Width           =   15855
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6600
         Width           =   1575
      End
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8040
         Width           =   1575
      End
      Begin VB.PictureBox picResults2 
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
         Height          =   2535
         Left            =   5760
         ScaleHeight     =   2475
         ScaleWidth      =   3675
         TabIndex        =   14
         Top             =   6480
         Width           =   3735
      End
      Begin VB.CommandButton cmdPurchaseForce 
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
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CommandButton cmdPurchase390 
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
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8160
         Width           =   1455
      End
      Begin VB.CommandButton cmdPurchaseRepublic 
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
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3600
         Width           =   1455
      End
      Begin VB.PictureBox picBinding3 
         Height          =   1695
         Left            =   1680
         Picture         =   "Form3.frx":34C0F
         ScaleHeight     =   1635
         ScaleWidth      =   1515
         TabIndex        =   4
         Top             =   7320
         Width           =   1575
      End
      Begin VB.PictureBox picBinding2 
         Height          =   1695
         Left            =   1680
         Picture         =   "Form3.frx":35FE8
         ScaleHeight     =   1635
         ScaleWidth      =   1515
         TabIndex        =   3
         Top             =   5040
         Width           =   1575
      End
      Begin VB.PictureBox picBinding1 
         Height          =   1695
         Left            =   1680
         Picture         =   "Form3.frx":370C9
         ScaleHeight     =   1635
         ScaleWidth      =   1515
         TabIndex        =   2
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "*Must Purchase Item Before Clicking Back To Ensure Correct Price!"
         Height          =   375
         Left            =   5400
         TabIndex        =   17
         Top             =   6120
         Width           =   5535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "$189.99"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   7680
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "$279.99"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   9
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "$219.99"
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
         Left            =   3480
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rome 390"
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
         Left            =   3480
         TabIndex        =   7
         Top             =   7320
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Force DLX"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Forum Republic"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   5
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pick Your Bindings"
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
         Height          =   1215
         Left            =   1680
         TabIndex        =   1
         Top             =   1560
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmThird"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPurchaseRepublic_Click()
Binding = 219.99
If (Money - Binding) < 0 Then
    MsgBox "Insufficient Funds... Please Go Back To Start!"
        picResults2.Cls
        frmThird.Hide
        frmFirst.Show
        cmdPurchaseRepublic.Enabled = True
        cmdPurchaseForce.Enabled = True
        cmdPurchase390.Enabled = True
Else
        cmdPurchaseRepublic.Enabled = False
        cmdPurchaseForce.Enabled = False
        cmdPurchase390.Enabled = False
End If
picResults2.Print " Items"; Tab(30); "Expenses"
picResults2.Print "*************************************************"
picResults2.Print Tab(30); FormatCurrency(Money)
picResults2.Print "Forum Republic"; Tab(30); FormatCurrency(-Binding)
picResults2.Print "*************************************************"
picResults2.Print "Total"; Tab(30); FormatCurrency(Money - Binding)

Money = Money - Binding

cmdNext.Enabled = True
cmdBack.Enabled = True
End Sub

Private Sub cmdPurchaseForce_Click()
Binding = 279.99
If (Money - Binding) < 0 Then
    MsgBox "Insufficient Funds... Please Go Back To Start!"
        picResults2.Cls
        frmThird.Hide
        frmFirst.Show
        cmdPurchaseRepublic.Enabled = True
        cmdPurchaseForce.Enabled = True
        cmdPurchase390.Enabled = True
Else
        cmdPurchaseRepublic.Enabled = False
        cmdPurchaseForce.Enabled = False
        cmdPurchase390.Enabled = False
End If
picResults2.Print " Items"; Tab(30); "Expenses"
picResults2.Print "*************************************************"
picResults2.Print Tab(30); FormatCurrency(Money)
picResults2.Print "Force DLX"; Tab(30); FormatCurrency(-Binding)
picResults2.Print "*************************************************"
picResults2.Print "Total"; Tab(30); FormatCurrency(Money - Binding)

Money = Money - Binding

cmdNext.Enabled = True
cmdBack.Enabled = True
End Sub

Private Sub cmdPurchase390_Click()
Binding = 189.99
If (Money - Binding) < 0 Then
    MsgBox "Insufficient Funds... Please Go Back To Start!"
        picResults2.Cls
        frmThird.Hide
        frmFirst.Show
        cmdPurchaseRepublic.Enabled = True
        cmdPurchaseForce.Enabled = True
        cmdPurchase390.Enabled = True
Else
        cmdPurchaseRepublic.Enabled = False
        cmdPurchaseForce.Enabled = False
        cmdPurchase390.Enabled = False
End If
picResults2.Print " Items"; Tab(30); "Expenses"
picResults2.Print "*************************************************"
picResults2.Print Tab(30); FormatCurrency(Money)
picResults2.Print "Rome 390"; Tab(30); FormatCurrency(-Binding)
picResults2.Print "*************************************************"
picResults2.Print "Total"; Tab(30); FormatCurrency(Money - Binding)

Money = Money - Binding

cmdNext.Enabled = True
cmdBack.Enabled = True
End Sub

Private Sub cmdBack_Click()
picResults2.Cls
Money = Money + Binding + Board
cmdPurchaseRepublic.Enabled = True
cmdPurchaseForce.Enabled = True
cmdPurchase390.Enabled = True

frmThird.Hide
frmSecond.Show
End Sub

Private Sub cmdNext_Click()
picResults2.Cls

cmdPurchaseRepublic.Enabled = True
cmdPurchaseForce.Enabled = True
cmdPurchase390.Enabled = True
cmdBack.Enabled = False
frmThird.Hide
frmFourth.Show

End Sub

Private Sub Form_Load()
cmdPurchaseRepublic.Enabled = True
cmdPurchaseForce.Enabled = True
cmdPurchase390.Enabled = True
End Sub
