VERSION 5.00
Begin VB.Form frmHistory 
   BackColor       =   &H8000000D&
   Caption         =   "History of Nascar"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLegends 
      Caption         =   "Legendary People of Nascar"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "How Nascar started"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   4845
      Left            =   7080
      Picture         =   "frmHistory.frx":0000
      Top             =   0
      Width           =   3555
   End
   Begin VB.Image Image1 
      Height          =   4725
      Left            =   120
      Picture         =   "frmHistory.frx":1362E
      Top             =   720
      Width           =   6000
   End
   Begin VB.Image Image2 
      Height          =   2415
      Left            =   5520
      Picture         =   "frmHistory.frx":1784A
      Top             =   4920
      Width           =   5655
   End
   Begin VB.Label lblHeader 
      Caption         =   "Nascar has changed a lot since it first started.  Let us get an idea of its legendary heritage."
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Introduction to NASCAR
'Form History
'Colin Roberts and Luke Hommerding
'Written 10/18/09
'Purpose is to have the history of nascar be displayed in a message box to the user
'and also leads to a form using dynamic picture loading
Option Explicit
    'when clicked history appears to the user
Private Sub cmdClick_Click()
    MsgBox ("Nascar was formed when stock cars were modified to transport moonshine during prohibition and be able to out run the law.  This started a stock car racing experience nationwide in the U.S. with no major organization to host a consistent season of competitive events.  This all changed when Bill France Senior of Daytona Beach held the first NASCAR race on the beach on Febuary 15, 1948.")
End Sub
    'takes users to a new form involving legends of NASCAR
Private Sub cmdLegends_Click()
    frmPeople.Show
    frmHistory.Hide
End Sub
    'returns user to main menu
Private Sub cmdReturn_Click()
    frmMain.Show
    frmHistory.Hide
End Sub

