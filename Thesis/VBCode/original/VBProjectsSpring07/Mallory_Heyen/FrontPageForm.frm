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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2880
      Picture         =   "FrontPageForm.frx":0000
      ScaleHeight     =   705
      ScaleWidth      =   3465
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.PictureBox picMen 
      Height          =   2655
      Left            =   5520
      Picture         =   "FrontPageForm.frx":88FA
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.PictureBox picWomen 
      Height          =   2655
      Left            =   1440
      Picture         =   "FrontPageForm.frx":1A25C
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   3840
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
      Left            =   3960
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdMen 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Click Here for Men's Fashions"
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
      Left            =   3120
      MaskColor       =   &H80000006&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblSign 
      BackColor       =   &H00000000&
      Caption         =   " Welcome to"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   3255
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



Private Sub cmdMen_Click()
    frmIntro.Visible = False
    frmMen.Visible = True
End Sub

Private Sub cmdWomen_Click()
    frmIntro.Visible = False
    frmWomen.Visible = True
    
    

End Sub

