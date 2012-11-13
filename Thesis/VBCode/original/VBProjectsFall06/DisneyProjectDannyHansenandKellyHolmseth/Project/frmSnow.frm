VERSION 5.00
Begin VB.Form frmSnow 
   BackColor       =   &H00FF0000&
   Caption         =   "Snow White and the Seven Dwarfs"
   ClientHeight    =   7560
   ClientLeft      =   2715
   ClientTop       =   1920
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   9690
   Begin VB.PictureBox Picture2 
      Height          =   4215
      Left            =   3600
      Picture         =   "frmSnow.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   5955
      TabIndex        =   3
      Top             =   3000
      Width           =   6015
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080C0FF&
      Caption         =   "Back"
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   480
      Picture         =   "frmSnow.frx":97A6
      ScaleHeight     =   3435
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   $"frmSnow.frx":C91D
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSnow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/30/06
'Objective: The objective of this form  is to display to the user a summary of the movie "Snow White and the Seven Dwarfs"
Private Sub cmdBack_Click()
frmSnow.Hide
frmTop.Show
End Sub
