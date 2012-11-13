VERSION 5.00
Begin VB.Form frmFourth 
   Caption         =   "Pick Your Boots"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   9135
      Left            =   0
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   9075
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   0
      Width           =   10935
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
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Next-->"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6240
         Width           =   1935
      End
      Begin VB.PictureBox picResults3 
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
         Height          =   2895
         Left            =   6360
         ScaleHeight     =   2835
         ScaleWidth      =   3915
         TabIndex        =   14
         Top             =   5280
         Width           =   3975
      End
      Begin VB.CommandButton cmdPurchasePrion 
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton cmdPurchaseID 
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton cmdPurchaseLashed 
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4440
         Width           =   1575
      End
      Begin VB.PictureBox Picture4 
         Height          =   1815
         Left            =   4080
         Picture         =   "Form4.frx":1C212
         ScaleHeight     =   1755
         ScaleWidth      =   1515
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.PictureBox picDeluxe 
         Height          =   1815
         Left            =   2280
         Picture         =   "Form4.frx":1D4DF
         ScaleHeight     =   1755
         ScaleWidth      =   1515
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin VB.PictureBox picLashed 
         Height          =   1815
         Left            =   480
         Picture         =   "Form4.frx":1E6DB
         ScaleHeight     =   1755
         ScaleWidth      =   1515
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "*Must Purchase Item Before Clicking Back To Ensure Correct Price!"
         Height          =   375
         Left            =   6000
         TabIndex        =   17
         Top             =   4920
         Width           =   4815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "$129.99"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4320
         TabIndex        =   10
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "$239.99"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "$199.99"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "32 Prion"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Deeluxe ID"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "32 Lashed"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pick Your Boots"
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
         Height          =   1095
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmFourth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub cmdPurchaseLashed_Click()
Boot = 199.99
If (Money - Boot) < 0 Then
    MsgBox "Insufficient Funds... Please Go Back To Start!"
        picResults3.Cls
        frmFourth.Hide
        frmFirst.Show
        cmdPurchaseLashed.Enabled = True
        cmdPurchaseID.Enabled = True
        cmdPurchasePrion.Enabled = True
Else
        picResults3.Cls
        cmdPurchaseLashed.Enabled = False
        cmdPurchaseID.Enabled = False
        cmdPurchasePrion.Enabled = False
End If

picResults3.Print " Items"; Tab(30); "Expenses"
picResults3.Print "*************************************************"
picResults3.Print Tab(30); FormatCurrency(Money)
picResults3.Print "32 Lashed"; Tab(30); FormatCurrency(-Boot)
picResults3.Print "*************************************************"
picResults3.Print "Total"; Tab(30); FormatCurrency(Money - Boot)

Money = Money - Boot

cmdNext.Enabled = True
cmdBack.Enabled = True
End Sub

Private Sub cmdPurchaseID_Click()
Boot = 239.99
If (Money - Boot) < 0 Then
    MsgBox "Insufficient Funds... Please Go Back To Start!"
        picResults3.Cls
        frmFourth.Hide
        frmFirst.Show
        cmdPurchaseLashed.Enabled = True
        cmdPurchaseID.Enabled = True
        cmdPurchasePrion.Enabled = True
Else
        picResults3.Cls
        cmdPurchaseLashed.Enabled = False
        cmdPurchaseID.Enabled = False
        cmdPurchasePrion.Enabled = False
End If
picResults3.Print " Items"; Tab(30); "Expenses"
picResults3.Print "*************************************************"
picResults3.Print Tab(30); FormatCurrency(Money)
picResults3.Print "Deeluxe ID"; Tab(30); FormatCurrency(-Boot)
picResults3.Print "*************************************************"
picResults3.Print "Total"; Tab(30); FormatCurrency(Money - Boot)

Money = Money - Boot

cmdNext.Enabled = True
cmdBack.Enabled = True
End Sub

Private Sub cmdPurchasePrion_Click()
Boot = 129.99
If (Money - Boot) < 0 Then
    MsgBox "Insufficient Funds... Please Go Back To Start!"
        picResults3.Cls
        frmFourth.Hide
        frmFirst.Show
        cmdPurchaseLashed.Enabled = True
        cmdPurchaseID.Enabled = True
        cmdPurchasePrion.Enabled = True
Else
        picResults3.Cls
        cmdPurchaseLashed.Enabled = False
        cmdPurchaseID.Enabled = False
        cmdPurchasePrion.Enabled = False
End If
picResults3.Print " Items"; Tab(30); "Expenses"
picResults3.Print "*************************************************"
picResults3.Print Tab(30); FormatCurrency(Money)
picResults3.Print "32 Prion"; Tab(30); FormatCurrency(-Boot)
picResults3.Print "*************************************************"
picResults3.Print "Total"; Tab(30); FormatCurrency(Money - Boot)

Money = Money - Boot

cmdNext.Enabled = True
cmdBack.Enabled = True
End Sub

Private Sub cmdBack_Click()
picResults3.Cls
Money = Money + Boot + Binding
cmdPurchaseLashed.Enabled = True
cmdPurchaseID.Enabled = True
cmdPurchasePrion.Enabled = True

frmFourth.Hide
frmThird.Show

End Sub
Private Sub cmdNext_Click()
picResults3.Cls

cmdPurchaseLashed.Enabled = True
cmdPurchaseID.Enabled = True
cmdPurchasePrion.Enabled = True
cmdBack.Enabled = False
frmFourth.Hide
frmFifth.Show
End Sub
Private Sub Form_Load()
cmdPurchaseLashed.Enabled = True
cmdPurchaseID.Enabled = True
cmdPurchasePrion.Enabled = True
cmdNext.Enabled = False
End Sub
