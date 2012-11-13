VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000D&
   Caption         =   "The Record -- Display or Classified"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form4"
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
      Left            =   8280
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   7920
      Picture         =   "Part2.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   2760
      Picture         =   "Part2.frx":7572
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   480
      Picture         =   "Part2.frx":1CBE4
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Displaybutton 
      Caption         =   "Click here to place a display ad."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   2880
      Width           =   3975
   End
   Begin VB.CommandButton Classified 
      Caption         =   "Click here to place a classified ad."
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Part2.frx":24156
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
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   1920
      Width           =   7575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "OR"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4800
      TabIndex        =   0
      Top             =   3720
      Width           =   375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1 (Record_Advertising), Form 4(Part2.frm), Kristen Nowak, 3-14-04, The purpose of this form is to allow the user to select whether or not they want to choose a classified or display ad.

Private Sub Classified_Click()
Form4.Hide 'If they choose classified, go to the classified form
Form7.Show
End Sub

Private Sub Displaybutton_Click()
Form5.Show 'If they choose display, go to the display form
Form4.Hide
End Sub

Private Sub Quitbutton_Click()
End
End Sub
