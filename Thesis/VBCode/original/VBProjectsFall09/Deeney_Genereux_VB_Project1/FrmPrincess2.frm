VERSION 5.00
Begin VB.Form FrmPrincess2 
   BackColor       =   &H0000FF00&
   Caption         =   "Princess2"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   14
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton CmdContinue 
      Caption         =   "Continue!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   13
      Top             =   3960
      Width           =   2655
   End
   Begin VB.PictureBox PicResults2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1920
      ScaleHeight     =   1275
      ScaleWidth      =   4755
      TabIndex        =   12
      Top             =   4200
      Width           =   4815
   End
   Begin VB.CommandButton CmdCalculate 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Calculate!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      MaskColor       =   &H00FFFF80&
      TabIndex        =   11
      Top             =   3120
      Width           =   5535
   End
   Begin VB.TextBox TxtY1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Text            =   "0"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox TxtX1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Text            =   "0"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox TxtY 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Text            =   "0"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox TxtX 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Text            =   "0"
      Top             =   1560
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   240
      Picture         =   "FrmPrincess2.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Lbl6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Y1="
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lbl5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X1="
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmPrincess2.frx":13ACA
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      TabIndex        =   6
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label Lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Y ="
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "X ="
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Now that you have picked out the perfect outfit for the princess, it is now time to test the princess' knowledge!"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9975
   End
End
Attribute VB_Name = "FrmPrincess2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'In this form, The princess finds the distance between two points
'The user inputs the values in text boxes, then the computer
'uses mathematical functions to solve.


Private Sub CmdCalculate_Click()
    Dim X As Integer, Y As Integer, X1 As Integer, Y1 As Integer, total As Single
    X = TxtX.Text
    Y = TxtY.Text
    X1 = TxtX1.Text
    Y1 = TxtY1.Text
    total = Sqr((X1 - X) * (X1 - X) + (Y1 - Y) * (Y1 - Y))
    PicResults2.Print "Too Bad you can't stump the princess "
    PicResults2.Print "because the Princess has come to an answer! "
    PicResults2.Print "She figured out that the distance between "
    PicResults2.Print "the two points is: "; FormatNumber(total, 2); ""
    PicResults2.Print "Now that you tested the princess' knowledge,"
    PicResults2.Print "You can now move on to the next part of the journey. "
    CmdContinue.Enabled = True
    
    

End Sub


Private Sub CmdContinue_Click()
    FrmPrincess2.Hide
    FrmPrincess3.Show
    PicResults2.Cls
    CmdContinue.Enabled = False
    
End Sub

Private Sub CmdQuit_Click()
    End
End Sub
