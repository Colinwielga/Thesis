VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00808080&
   Caption         =   "Intro"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWomen2 
      Height          =   2655
      Left            =   2520
      Picture         =   "frmIntro.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   3720
      Width           =   2055
   End
   Begin VB.PictureBox picBrownBag 
      Height          =   2895
      Left            =   5640
      Picture         =   "frmIntro.frx":11962
      ScaleHeight     =   2835
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2880
      Picture         =   "frmIntro.frx":39064
      ScaleHeight     =   705
      ScaleWidth      =   3465
      TabIndex        =   3
      Top             =   840
      Width           =   3495
   End
   Begin VB.PictureBox picWomen 
      Height          =   2655
      Left            =   240
      Picture         =   "frmIntro.frx":4195E
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdWomen 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      Caption         =   "Click Here for Women'S Fashions"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2640
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdMen 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Click Here for Special discounting"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5760
      MaskColor       =   &H80000006&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome To"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program will let the user view different pieces of clothing
'that the store offers and then choose to add them to their purchase
'or look at a different piece of clothing.  The user is then able to
'proceed to the checkout and see if any of the current discounts
'apply to their purchase.


'Move from this form to the discount form using the visible variable
Private Sub cmdMen_Click()
    frmIntro.Visible = False
    frmDiscount.Visible = True
End Sub
'Move from this form to the womens form using the visible variable
Private Sub cmdWomen_Click()
    frmIntro.Visible = False
    frmWomen.Visible = True
    
End Sub


