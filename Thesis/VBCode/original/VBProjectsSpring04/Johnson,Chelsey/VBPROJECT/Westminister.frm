VERSION 5.00
Begin VB.Form Westminister 
   BackColor       =   &H00FF00FF&
   Caption         =   "Westminister"
   ClientHeight    =   14850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14160
      TabIndex        =   13
      Top             =   12840
      Width           =   1695
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11520
      TabIndex        =   12
      Top             =   12720
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Text            =   "Click on either picture to learn more "
      Top             =   7680
      Width           =   3255
   End
   Begin VB.PictureBox piceye2 
      Height          =   7815
      Left            =   11400
      Picture         =   "Westminister.frx":0000
      ScaleHeight     =   7755
      ScaleWidth      =   4395
      TabIndex        =   9
      Top             =   3600
      Width           =   4455
   End
   Begin VB.PictureBox piceye 
      Height          =   2295
      Left            =   8760
      Picture         =   "Westminister.frx":95A5
      ScaleHeight     =   2235
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Text            =   "Click on either picture to learn about the history of big ben."
      Top             =   6840
      Width           =   5295
   End
   Begin VB.PictureBox picBen2 
      Height          =   5415
      Left            =   360
      Picture         =   "Westminister.frx":DCD1
      ScaleHeight     =   5355
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   6120
      Width           =   1695
   End
   Begin VB.PictureBox picBen 
      Height          =   3015
      Left            =   3480
      Picture         =   "Westminister.frx":11BBC
      ScaleHeight     =   2955
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   7560
      Width           =   2055
   End
   Begin VB.PictureBox picParliment 
      Height          =   2535
      Left            =   7440
      Picture         =   "Westminister.frx":133C1
      ScaleHeight     =   2475
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
   Begin VB.Label Label6 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   14400
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "London Eye at night and day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   10
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "These pictures are both of Big Ben"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   6240
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "   Parliament"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Click on the pictures to learn more about each one"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Famous Sites of Westminister District in London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "Westminister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: Westminister (Westminister.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: The purpose of this form was to let the user click on a picture of a site of their choice
                    'and then have a message box appear with the history of the site
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'User returns to Map of London page to choose a new district
Westminister.Hide
MapLondon.Show
End Sub

Private Sub picBen_Click()
'The history of Big Ben is given in the form of a Message Box
MsgBox "Big Ben is one of the most famous landmarks in the world. The clock tower situated on the banks of the river Thames epitomises the culture and architectural style of London and has tolled out the hours before the news on BBC radio since 1923. Officially Big Ben is only the name of the biggest of the five bells in the clock tower also known as St Stephen's Tower. The 13.8 ton bell was cast at Whitechapel Bell Foundry in 1858 and is said to have been pulled to the site by 16 horses. It was installed in 1859, but did not start ringing until 1862 because of a crack.", , "Big Ben"
MsgBox "The tower itself is 96.3m/316ft high. There are 4 clock faces, which have a diameter of 7.5m each. The minute hands are made of copper and are 4.3m long. ", , "Big Ben"
End Sub

Private Sub picBen2_Click()
'The history of Big Ben is given in the form of a Message Box
MsgBox "Big Ben is one of the most famous landmarks in the world. The clock tower situated on the banks of the river Thames epitomises the culture and architectural style of London and has tolled out the hours before the news on BBC radio since 1923. Officially Big Ben is only the name of the biggest of the five bells in the clock tower also known as St Stephen's Tower. The 13.8 ton bell was cast at Whitechapel Bell Foundry in 1858 and is said to have been pulled to the site by 16 horses. It was installed in 1859, but did not start ringing until 1862 because of a crack.", , "Big Ben"
MsgBox "The tower itself is 96.3m/316ft high. There are 4 clock faces, which have a diameter of 7.5m each. The minute hands are made of copper and are 4.3m long. ", , "Big Ben"
End Sub

Private Sub piceye_Click()
'The history of The London Eye is given in the form of a Message Box
MsgBox "The London eye is the worlds largest observation wheel.", , "London Eye"
MsgBox "The London eye was originally called the Millennium Wheel. The London Eye represents the turning of time and the cycle of life. The London Eye was built horizontally and then lifted by a crane into the vertical position.", , "London Eye"
End Sub

Private Sub piceye2_Click()
'The history of The London Eye is given in the form of a Message Box
MsgBox "The London eye is the worlds largest observation wheel.", , "London Eye"
MsgBox "The London eye was originally called the Millennium Wheel. The London Eye represents the turning of time and the cycle of life. The London Eye was built horizontally and then lifted by a crane into the vertical position.", , "London Eye"
End Sub

Private Sub picParliment_Click()
'The history of The Parliament is given in the form of a Message Box
MsgBox "The UK Parliament is based on a two chamber system. The House of Lords and the House of Commons sit separately, and are constituted on different principles. However, the legislative process involves both Houses. Parliament has three main functions- 1. to examine proposals for new laws, 2. to scrutinise government policy and administration, and 3. to debate the major issues of the day.", , "Parliament"
MsgBox "Parliament has gradually taken control over many of the powers previously exercised by the Monarch. The Monarch now has a constitutional role which means that their actions are governed by convention.", , "Parliament"
End Sub
