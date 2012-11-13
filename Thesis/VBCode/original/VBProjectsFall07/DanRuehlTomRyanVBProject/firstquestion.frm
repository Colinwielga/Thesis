VERSION 5.00
Begin VB.Form frmQuestion 
   BackColor       =   &H00800000&
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtAnswer1 
      Height          =   615
      Left            =   11880
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAnswer 
      Height          =   615
      Left            =   11880
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   4335
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   4335
   End
   Begin VB.CommandButton cmdB 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   4335
   End
   Begin VB.CommandButton cmdA 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   120
      Picture         =   "firstquestion.frx":0000
      Top             =   4080
      Width           =   12885
   End
   Begin VB.Label lblQuestions 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9495
   End
End
Attribute VB_Name = "frmQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTRq As Integer


Private Sub cmdA_Click()
frmJeopardy.picresults1.Cls
txtAnswer1.Text = "A"
    
        If txtAnswer.Text = txtAnswer1.Text Then
            Sum = Sum + F
            MsgBox ("Correct!")
        Else
            Sum = Sum - F
            MsgBox ("Wrong!")
        End If
        CTRq = CTRq + 1
        If CTRq = 30 Then
            frmJeopardy.cmdFinalJeopardy.Visible = True
        End If
        lblQuestions.Caption = ""
        cmdA.Caption = ""
        cmdB.Caption = ""
        cmdC.Caption = ""
        cmdD.Caption = ""
        frmJeopardy.Show
        frmQuestion.Hide
        frmJeopardy.picResults.Print "$"; FormatNumber(Sum)
        frmJeopardy.picresults1.Print "$"; FormatNumber(Sum)
End Sub

Private Sub cmdB_Click()
frmJeopardy.picresults1.Cls
 txtAnswer1.Text = "B"
    
        If txtAnswer.Text = txtAnswer1.Text Then
            Sum = Sum + F
            MsgBox ("Correct!")
        Else
            Sum = Sum - F
            MsgBox ("Wrong!")
        End If
        CTRq = CTRq + 1
        If CTRq = 30 Then
            frmJeopardy.cmdFinalJeopardy.Visible = True
        End If
        lblQuestions.Caption = ""
        cmdA.Caption = ""
        cmdB.Caption = ""
        cmdC.Caption = ""
        cmdD.Caption = ""
        frmJeopardy.Show
        frmQuestion.Hide
        frmJeopardy.picResults.Print "$"; FormatNumber(Sum)
        frmJeopardy.picresults1.Print "$"; FormatNumber(Sum)
End Sub


Private Sub cmdC_Click()
frmJeopardy.picresults1.Cls
 txtAnswer1.Text = "C"
    
        If txtAnswer.Text = txtAnswer1.Text Then
            Sum = Sum + F
            MsgBox ("Correct!")
        Else
            Sum = Sum - F
            MsgBox ("Wrong!")
        End If
        CTRq = CTRq + 1
        If CTRq = 30 Then
            frmJeopardy.cmdFinalJeopardy.Visible = True
        End If
        lblQuestions.Caption = ""
        cmdA.Caption = ""
        cmdB.Caption = ""
        cmdC.Caption = ""
        cmdD.Caption = ""
        frmJeopardy.Show
        frmQuestion.Hide
        frmJeopardy.picResults.Print "$"; FormatNumber(Sum)
        frmJeopardy.picresults1.Print "$"; FormatNumber(Sum)
End Sub

Private Sub cmdD_Click()
frmJeopardy.picresults1.Cls
 txtAnswer1.Text = "D"
    
        If txtAnswer.Text = txtAnswer1.Text Then
            Sum = Sum + F
            MsgBox ("Correct!")
        Else
            Sum = Sum - F
            MsgBox ("Wrong!")
        End If
        CTRq = CTRq + 1
        If CTRq = 30 Then
            frmJeopardy.cmdFinalJeopardy.Visible = True
        End If
        lblQuestions.Caption = ""
        cmdA.Caption = ""
        cmdB.Caption = ""
        cmdC.Caption = ""
        cmdD.Caption = ""
        frmJeopardy.Show
        frmQuestion.Hide
        frmJeopardy.picResults.Print "$"; FormatNumber(Sum)
        frmJeopardy.picresults1.Print "$"; FormatNumber(Sum)
End Sub


