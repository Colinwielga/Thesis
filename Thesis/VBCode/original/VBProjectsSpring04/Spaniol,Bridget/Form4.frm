VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000002&
   Caption         =   "Form4"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14655
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form4"
   ScaleHeight     =   8850
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdshow 
      BackColor       =   &H8000000E&
      Caption         =   "CLICK HERE to learn more about International Paper's financial stance through the interpretation of their financial statements."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bridget Spaniol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12960
      TabIndex        =   3
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INTERNATIONAL PAPER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   9735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   $"Form4.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   2760
      Picture         =   "Form4.frx":00C3
      Top             =   2040
      Width           =   9000
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'International Paper financial interpretation (Financial Analysis)
'Form 4 (project 1)
'Bridget Spaniol
'3/13/04
'The purpose of this program is to inform the user of the financial status of International Paper
'through the interpretation of the company's income and balance sheets.
'This form is the introduction form for the user to gain a little background about the company and what they will
'be doing with the program.

Private Sub cmdshow_Click()
Form1.Show
Form4.Hide
End Sub

