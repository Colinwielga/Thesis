VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H8000000D&
   Caption         =   "The Record -- Display ad -- color or black & white"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form5"
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
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   7920
      Picture         =   "Display.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   480
      Picture         =   "Display.frx":7572
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   2760
      Picture         =   "Display.frx":EAE4
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Color 
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton Black 
      Caption         =   "Black and White"
      BeginProperty Font 
         Name            =   "Batang"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Would you like your ad to be in color or black and white?"
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
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2040
      Width           =   2895
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1 (Record_Advertising), Form 5(display.frm), Kristen Nowak, 3-14-04, The purpose of this form is to allow the user to select whether they want a black & white or color ad.

Private Sub Black_Click()
Form5.Hide 'If they choose black & white, go to the black & white form
Form3.Show
End Sub

Private Sub Color_Click()
Form5.Hide 'If they choose color, go to the color form
Form6.Show
End Sub

Private Sub Quitbutton_Click()
End
End Sub
