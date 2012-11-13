VERSION 5.00
Begin VB.Form frmWomen 
   BackColor       =   &H00400040&
   Caption         =   "Women"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHandbags 
      Height          =   2655
      Left            =   3960
      Picture         =   "fmrWomen.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.PictureBox picShoes 
      Height          =   2655
      Left            =   3960
      Picture         =   "fmrWomen.frx":11962
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   5040
      Width           =   2055
   End
   Begin VB.PictureBox picDress 
      Height          =   2655
      Left            =   6240
      Picture         =   "fmrWomen.frx":232C4
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.PictureBox picJeans 
      Height          =   2655
      Left            =   6240
      Picture         =   "fmrWomen.frx":3507E
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdChoices 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Here to Begin Shopping"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblChoices 
      Caption         =   "   Dresses                       Handbags                        Jeans                           Shoes"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   960
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblWomen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to Women's Apparel"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmWomen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Form Objectices: Ask the user to input what type of apparel they
'would like to view based on a list included on the form.  Based
'on the input the program will bring the user to the form the customer
'chose.

'Define All Variables

Dim Apparel As String
Dim Dresses As String
Dim Shoes As String
Dim Handbags As String
Dim Jeans As String

Private Sub cmdChoices_Click()
'The input box will ask the user to specify what kind of apparel
'they would like to see based on the options on the form
Apparel = InputBox("What type of apparel from the list would you like to see?", "?")
    'The Select Case option will compare the user's input with the
    'different form option and bring the user to the correct form
    'based on the user's desired apparel
    Select Case Apparel
        Case Is = "Dresses"
            frmWomen.Visible = False
            frmDresses.Visible = True
        Case Is = "Shoes"
            frmWomen.Visible = False
            frmShoes.Visible = True
        Case Is = "Handbags"
            frmWomen.Visible = False
            frmHandbags.Visible = True
        Case Is = "Jeans"
            frmWomen.Visible = False
            frmJeans.Visible = True
    End Select
    
End Sub


