VERSION 5.00
Begin VB.Form frmForm1 
   Caption         =   "Study Abroad"
   ClientHeight    =   6450
   ClientLeft      =   3930
   ClientTop       =   2325
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   Picture         =   "frmForm1.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   7950
   Begin VB.CommandButton cmdGetStarted 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Here to Get Started"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblStudyAbroad 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Study Abroad Assistant"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   54.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   7935
   End
End
Attribute VB_Name = "frmForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'starting form
'written 3/9/08 by Sammi and Erika


Private Sub cmdGetStarted_Click()
'brings user to next form where they can choose programs to look at

    frmForm1.Hide
    frmPrograms.Show
    
End Sub

Private Sub cmdSlideShow_Click()
    frmForm1.Hide
    frmPictures.Show
End Sub
