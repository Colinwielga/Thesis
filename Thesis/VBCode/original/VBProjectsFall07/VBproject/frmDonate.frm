VERSION 5.00
Begin VB.Form frmDonate 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWow 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Thelma After Donations!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      Picture         =   "frmDonate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CommandButton cmdhome 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Go Home!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5520
      Picture         =   "frmDonate.frx":0A4E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   6015
      Left            =   7920
      ScaleHeight     =   5955
      ScaleWidth      =   3795
      TabIndex        =   7
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5520
      MaskColor       =   &H8000000F&
      Picture         =   "frmDonate.frx":1887
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdcalculate 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Calculate!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5640
      Picture         =   "frmDonate.frx":1F6C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdboyfriend 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Boyfriend"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3240
      Picture         =   "frmDonate.frx":2CB1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdhair 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Hair Stylist"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      Picture         =   "frmDonate.frx":3789
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdmakeup 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Makeup"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2880
      Picture         =   "frmDonate.frx":438E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdclothes 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Clothes"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      Picture         =   "frmDonate.frx":4EEF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label lblhelp 
      Caption         =   "Donate To Thelma's Cause:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmDonate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningtotal As Single


Private Sub cmdclothes_Click()
runningtotal = runningtotal + 50
picResults.Print "New Clothes ", FormatCurrency(50#)
End Sub

Private Sub cmdClear_Click()
picResults.Cls
runningtotal = 0
End Sub

Private Sub cmdcalculate_Click()
Dim Tax As Single
Dim Subtotal As Single
picResults.Print "-----------------------------------------"
picResults.Print "Subtotal ", FormatCurrency(runningtotal)
Tax = runningtotal * 0.07
picResults.Print "Tax ", FormatCurrency(Tax)
total = runningtotal + Tax
picResults.Print "Total ", FormatCurrency(total)
End Sub

Private Sub cmdHome_Click()
frmDonate.Hide
frmDoll.Show

End Sub

Private Sub cmdmakeup_Click()
'adds amount to runningtotal and prints what was purchased
runningtotal = runningtotal + 35#
picResults.Print "Make-up ", FormatCurrency(35)
End Sub

Private Sub cmdQuit_Click()
'Exits program
End
End Sub

Private Sub cmdhair_Click()
'adds amount to runningtotal and prints what was purchased
runningtotal = runningtotal + 75
picResults.Print "Hair Stylist ", FormatCurrency(75)

End Sub

Private Sub cmdboyfriend_Click()
'Displays a Msgbox
MsgBox ("Boyfriend = PRICELESS!")
picResults.Print "Boyfriend ", "Priceless"
End Sub




Private Sub cmdWow_Click()
'Opens new form
frmDonate.Hide
frmPretty.Show


 
End Sub
