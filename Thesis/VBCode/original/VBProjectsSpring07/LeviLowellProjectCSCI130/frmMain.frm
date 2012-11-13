VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Main Menu"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C000&
      Caption         =   "Close The Bank"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   3735
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H0000C000&
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdMeetme 
      BackColor       =   &H0000C000&
      Caption         =   "Meet the Creator"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdShop 
      BackColor       =   &H0000C000&
      Caption         =   "Shop for Lowell Merchandise"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdCurrencySorter 
      BackColor       =   &H0000C000&
      Caption         =   "Currency Sorter"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdBankAccount 
      BackColor       =   &H0000C000&
      Caption         =   "Your Bank Account"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdCurrency 
      BackColor       =   &H0000C000&
      Caption         =   "Currency Converter"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Created By: Levi Lowell            CSCI-130"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   7680
      Width           =   5895
   End
   Begin VB.Image Image2 
      Height          =   3375
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   7320
      Picture         =   "frmMain.frx":C2FD
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Image ImgMain 
      Height          =   4890
      Left            =   3360
      Picture         =   "frmMain.frx":185FA
      Top             =   3000
      Width           =   3690
   End
   Begin VB.Label lblBankIntro 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome to the ""all-in-one"" Lowell Bank!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2295
      Left            =   3360
      TabIndex        =   6
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Welcome to the Lowell Bank!  This is the main page.  This code contains the navigation
'and login data.  Once the user inputs their name and password they are capable of entering
'the other applications.  Otherwise this page of code provides program with the ability
'to navigate to other forms and programs.

Private Sub cmdBankAccount_Click()

If pos = False Then
        frmBankAccount.Hide
        MsgBox "You must login in order to view this page.  Please login in now.", , "Login Error!"
        FrmMain.Show
Else
        frmBankAccount.Show     'Shows frmBankAccount
        FrmMain.Hide        'Hides frmMain
End If
End Sub

Private Sub cmdCurrency_Click()

If pos = False Then
        frmCurrency.Hide
        MsgBox "You must login in order to view this page.  Please login in now.", , "Login Error!"
        FrmMain.Show
Else
       
        frmCurrency.Show        'Show frmCurrency
        FrmMain.Hide        'Hides frmMain
End If
End Sub


Private Sub cmdCurrencySorter_Click()

If pos = False Then
        frmSorter.Hide
        MsgBox "You must login in order to view this page.  Please login in now.", , "Login Error!"
        FrmMain.Show
Else
        frmSorter.Show      'Shows frmSorter
        FrmMain.Hide        'Hides frmMain
End If
End Sub

Private Sub cmdLogin_Click()
frmLogin.Show
FrmMain.Hide
End Sub

Private Sub cmdShop_Click()

If pos = False Then
        frmPurchases.Hide
        MsgBox "You must login in order to view this page.  Please login in now.", , "Login Error!"
        FrmMain.Show
Else
        frmPurchases.Show       'Shows frmPurchases
        FrmMain.Hide        'Hides frmMain
End If
End Sub

Private Sub CmdMeetme_Click()

If pos = False Then
        frmMeetme.Hide
        MsgBox "You must login in order to view this page.  Please login in now.", , "Login Error!"
        FrmMain.Show
Else
        frmMeetme.Show      'Shows frm Meetme
        FrmMain.Hide        'Hides frmMain
End If
End Sub

Private Sub CmdQuit_Click()
End     'Ends Program
End Sub

Private Sub Form_Load()
username = "Levi Lowell"
Password = "Soccer12"
End Sub
