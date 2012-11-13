VERSION 5.00
Begin VB.Form frmTax1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdmin 
      BackColor       =   &H000000FF&
      Caption         =   "Administration"
      BeginProperty Font 
         Name            =   "Gentium"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H00FF0000&
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   4320
      Picture         =   "frmTax1.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00000000&
      Caption         =   "Login: "
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00000000&
      Caption         =   "Welcome To Tax Genie 2007 !"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   4335
   End
End
Attribute VB_Name = "frmTax1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdmin_Click()
    frmTax1.Hide
    frmAdmin.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub


Private Sub cmdInput_Click()
    'this will require the user to input their name
    N = txtName.Text
    
    'This will open and read the data file into an array
    Open App.Path & "\Bracket.txt" For Input As #1
    
    CTR1 = 0
    
    Do Until EOF(1)
        CTR1 = CTR1 + 1
        Input #1, Risk(CTR1), Bracket(CTR1), Potential(CTR1)
    Loop
    Close #1
    
    frmTax1.Hide
    frmType.Show
    
End Sub

