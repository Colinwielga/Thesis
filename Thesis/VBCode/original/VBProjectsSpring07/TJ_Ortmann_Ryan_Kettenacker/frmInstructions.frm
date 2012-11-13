VERSION 5.00
Begin VB.Form frmInstructions 
   BackColor       =   &H000080FF&
   Caption         =   "The Rules and Scoring"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoToBrackets 
      Caption         =   "CLICK HERE TO START THE MADNESS!!"
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   7920
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHAMPIONSHIP:    12 POINTS FOR EVERY CORRECT PICK"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   5160
      Width           =   8175
   End
   Begin VB.Image Image2 
      Height          =   2970
      Left            =   6960
      Picture         =   "frmInstructions.frx":0000
      Top             =   120
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   2970
      Left            =   120
      Picture         =   "frmInstructions.frx":DC3B
      Top             =   120
      Width           =   3000
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmInstructions.frx":1B876
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2040
      TabIndex        =   6
      Top             =   6120
      Width           =   6255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHAMPION:    20 POINTS FOR PICKING THE WINNER!"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   5640
      Width           =   8175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "FINAL FOUR:    8 POINTS FOR EVERY CORRECT PICK"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   4680
      Width           =   8175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "ELITE EIGHT:    4 POINTS FOR EVERY CORRECT PICK"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   4200
      Width           =   8175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "SWEET SIXTEEN:    2 POINTS FOR EVERY CORRECT PICK"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   3720
      Width           =   8175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "ROUND 1:    1 POINTS FOR EVERY CORRECT PICK"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   3240
      Width           =   8175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "SCORING IS DONE EACH ROUND"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3360
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this form was to give the user rules and scoring of our bracket
'the button below is to go from the instruction form to the midwest bracket
Private Sub cmdGoToBrackets_Click()
    frmInstructions.Hide
    frmMidwest.Show
End Sub
