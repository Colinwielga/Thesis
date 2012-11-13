VERSION 5.00
Begin VB.Form frmCentralBankWelcome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Welcome to Central Bank"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   Picture         =   "CentralBank.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   6960
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Leave Bank"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdBankTeller 
      BackColor       =   &H00008000&
      Caption         =   "Continue To Bank Teller"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   4095
   End
   Begin VB.Image imaLogo 
      Height          =   4095
      Left            =   4680
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Loading..."
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The current date and time  is:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"CentralBank.frx":2D30A
      Height          =   1815
      Left            =   480
      TabIndex        =   2
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label lblWelcome2 
      BackColor       =   &H00FFFFFF&
      Caption         =   """We make it happen!"""
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   15960
      TabIndex        =   1
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome To Central Bank"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmCentralBankWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This bank system was designed and created by Mark Brown and David Bernardy

Option Explicit


Private Sub cmdBankTeller_Click()
frmCentralBankWelcome.Hide                              'Makes the form frmCentralBankWelcome hidden
frmCustomerChoose.Show                                  'Makes the form frmCustomerChoose visible

End Sub

Private Sub cmdQuit_Click()
End                                                     'Exits the bank

End Sub


Private Sub Form_Load()

imaLogo.Picture = LoadPicture(App.Path & "\logo.gif")   'Imports our bank logo


End Sub

Private Sub Timer1_Timer()

lblTime.Caption = Now                                   'This is the code for our clock

End Sub
