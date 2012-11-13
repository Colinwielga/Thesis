VERSION 5.00
Begin VB.Form frmPicturesJohn 
   BackColor       =   &H00C0FFFF&
   Caption         =   "John"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   14955
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   16680
      TabIndex        =   2
      Top             =   11400
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Click to go to the Previous Picture"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   16680
      TabIndex        =   1
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label lblJohn 
      Alignment       =   2  'Center
      Caption         =   "John"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   0
      Top             =   12000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   9360
      Left            =   3600
      Picture         =   "PicturesJohn.frx":0000
      Top             =   2280
      Width           =   12450
   End
End
Attribute VB_Name = "frmPicturesJohn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmPicturesJohn(PicturesJohn.frm)
'Sarah Keating
'10-27-03
'Purpose: This final form shows a picture of John.



Private Sub cmdPrevious_Click()
frmPicturesJason.Hide
frmPicturesPat.Show
' Allows the user to go to the previous photo
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub

