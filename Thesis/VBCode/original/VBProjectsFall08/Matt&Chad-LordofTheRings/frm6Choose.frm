VERSION 5.00
Begin VB.Form frm6Choose 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Pristina"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   Picture         =   "frm6Choose.frx":0000
   ScaleHeight     =   7575
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End Your Journey"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8760
      TabIndex        =   4
      Top             =   6960
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7440
      ScaleHeight     =   2115
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton CommancmdChoose 
      Caption         =   "Choose"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   7800
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm6Choose.frx":C894
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1815
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sam Merry Pippin Aragorn Gandalf Boromir Legolas Gimli"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frm6Choose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdContinue_Click()
    frm6Choose.Hide
    frm7.Show
End Sub

Private Sub CommancmdChoose_Click()
Dim MemberCTR As Integer, n As Integer, MemberName As String
    MsgBox ("You get to select 4 other companions, Frodo- choose wisely")
        
        Do Until MemberCTR = 4
            MemberCTR = MemberCTR + 1
            CTRmember = CTRmember + 1
            MemberName = InputBox("Please enter the name of your new member", "Choose your Fellowship")
            Member(CTRmember) = MemberName
        Loop
    MsgBox ("Congragulations you are ready to begin your quest to destroy the ring. You will be given 4 pieces of Lambis Bread to eat on the long road to Mount Doom")
    
For n = 1 To CTRmember
picResults.Print Member(n)
Next n
End Sub

Private Sub Command1_Click()
    frm6Choose.Hide
    frm2Characters.Show
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    Lambis = 4
End Sub

