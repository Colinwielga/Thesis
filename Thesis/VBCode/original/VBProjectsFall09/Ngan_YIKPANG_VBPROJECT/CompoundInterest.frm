VERSION 5.00
Begin VB.Form CompoundInterest 
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "CompoundInterest.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd3 
      Caption         =   "Go to next page to see some examples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   5
      Top             =   4800
      Width           =   2415
   End
   Begin VB.PictureBox Picresults 
      Height          =   5655
      Left            =   4680
      ScaleHeight     =   5595
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Formulas you can use to calculate "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   1
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H8000000B&
      Caption         =   "What is compound interest?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblCompoundInterest 
      BackColor       =   &H80000012&
      Caption         =   "Compound Interest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "CompoundInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Compound Interest
'Form:CompoundInterest
'Author:Yik Pang Ngan (Banny)
'Date Written:Oct 9 2009
Option Explicit
'this program will teach what is compound interest. formulas of how it works and examples of how to calculate.

Private Sub Label1_Click()

End Sub

Private Sub cmd1_Click()
Dim definition As String
'this buttom will show the definition of what is compound interest in the picture box

    Open App.Path & "\whatiscompoundinterest.txt" For Input As #1 'open the file from notepad
    
    Picresults.Print "Definition" 'print definition
    
    Do While Not EOF(1)
        Input #1, definition
        Picresults.Print definition
Loop
End Sub

Private Sub cmd2_Click()
CompoundInterest.Hide
Formulas.Show
'this buttom will switch to the next page to formulas

End Sub

Private Sub cmd3_Click()
CompoundInterest.Hide
example.Show
'this buttom will switch to the next page to example
End Sub

Private Sub cmdQuit_Click()
End

End Sub

