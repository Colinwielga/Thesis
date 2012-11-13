VERSION 5.00
Begin VB.Form Introduction 
   BackColor       =   &H00400040&
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitchForm 
      BackColor       =   &H0000FF00&
      Caption         =   "Experiment with financial statements"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   3735
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H008080FF&
      Caption         =   "More about Accounting"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   3735
   End
   Begin VB.PictureBox picResults 
      Height          =   4455
      Left            =   7920
      ScaleHeight     =   4395
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   3480
      Width           =   5415
   End
   Begin VB.CommandButton cmdBasic 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Basis of Accounting"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label lblIntro 
      BackColor       =   &H00FFC0FF&
      Caption         =   "                 Introduction"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "Introduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Accounting basics and Income statement
'Form 1:Introduction
'Author:Patrick Niyibizi and Frankie Chan
'Date Written:September 30th 2009
'Our purpose in doing this project is to introduce the principles of Accounting and financial statements. The user will learn by filling his/her own financial statements.
Option Explicit
Private Sub cmdBasic_Click()
    picResults.Picture = LoadPicture(App.Path & "\Images\basicAccounting.jpg")  'Display a picture that illustrates the basis of accounting.
End Sub

Private Sub cmdInfo_Click()
    MsgBox ("A company should prepare financial statements at the end of every accounting operating cycle." & vbNewLine & "This time period may vary, but in most cases, it would be at the end of the year, December 31.")
    MsgBox ("Cash is KING." & vbNewLine & "A huge net income does not necessarily mean a lot of cash." & vbNewLine & "This is because net income may come from receivables (sales on credit without cash)." & vbNewLine & "Cash is needed to pay loans, to pay expenses, and to keep a company from going bankrupt.")
    MsgBox ("Financial accounting involves general accounting entries, journal set up, and preparing financial statements such as trial balance, income statement, balance sheet, and cash-flow statement." & vbNewLine & "In the US, companies prepare these statements based on GAAP (General Accepted Accounting Policies).")
End Sub

Private Sub cmdSwitchForm_Click()     'Go to the next form and start experimenting with some numbers.
    Introduction.Hide
    experiment.Show
End Sub



