VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "The Record advertising department"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quitbutton 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   7920
      Picture         =   "adrates.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   480
      Picture         =   "adrates.frx":7572
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton BeginButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To calculate the cost of an ad, click here."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   3480
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   2760
      Picture         =   "adrates.frx":EAE4
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"adrates.frx":24156
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1 (Record_Advertising), Form 1(adrates.frm), Kristen Nowak, 3-14-04, The purpose of this form is to introduce the program

Private Sub Label1_Click()

End Sub

Private Sub BeginButton_Click()
Form4.Show 'open up the next page in the program
Form1.Hide 'hide the first page
End Sub

Private Sub Quitbutton_Click()
End 'quit
End Sub
