VERSION 5.00
Begin VB.Form frmBegin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Build"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   1800
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   9
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdSubcompact 
      Caption         =   "Subcompact"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCompact 
      Caption         =   "Compact"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCoupe 
      Caption         =   "Coupe"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdTruck 
      Caption         =   "Truck"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdSedan 
      Caption         =   "Sedan"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdSUV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SUV"
      Height          =   495
      Left            =   240
      Picture         =   "frmBegin.frx":0000
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblPick 
      Alignment       =   2  'Center
      Caption         =   "Pick a body style to build on."
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmBegin
'Author is Dan Parker
'Date written 10/12/09
'The purpose of this form is to provide a base for the user to choose
'which vehicle they wish to build.

Private Sub cmdBack_Click()
    frmBegin.Hide 'hide begin form
    frmFirst.Show 'show first form
End Sub

Private Sub cmdCompact_Click()
    
    frmBegin.Hide 'hide begin form
    frmBuildCompact.Show 'show compact car form
    
End Sub

Private Sub cmdCoupe_Click()
    
    frmBegin.Hide 'hide begin form
    frmBuildCoupe.Show 'show coupe form
    
End Sub

Private Sub cmdSedan_Click()
    
    frmBegin.Hide 'hide begin form
    frmBuildSedan.Show 'show sedan form

End Sub

Private Sub cmdSubcompact_Click()

    frmBegin.Hide 'hide begin form
    frmBuildSubcompact.Show 'show subcompact form
    
End Sub

Private Sub cmdSUV_Click()
    
    frmBegin.Hide 'hide begin form
    frmBuildSUV.Show 'show SUV form
    
End Sub

Private Sub cmdTruck_Click()
    
    frmBegin.Hide 'hide begin form
    frmBuildTruck.Show 'show truck form
    
End Sub


Private Sub cmdQuit_Click()
MsgBox ("Thanks for using the Wicked Fun Car Builder, " & " " & UserName & "!")
End 'end program
End Sub

Private Sub Form_Load()
    'load picture onto form
    picResults.Picture = LoadPicture(App.Path & "\bmw.jpg")
End Sub
