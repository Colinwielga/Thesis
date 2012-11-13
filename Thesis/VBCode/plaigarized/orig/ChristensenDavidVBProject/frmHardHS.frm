VERSION 5.00
Begin VB.Form frmHardHS 
   BackColor       =   &H00000000&
   Caption         =   "High Scores"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   2175
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   495
      Left            =   3360
      TabIndex        =   21
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   855
      Left            =   3360
      TabIndex        =   20
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txt5 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Text            =   "999.9"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txt4 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Text            =   "999.9"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txt3 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Text            =   "999.9"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "999.9"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "999.9"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtname1 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtname2 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtname3 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtname4 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtname5 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00000000&
      Caption         =   "1."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00000000&
      Caption         =   "2."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00000000&
      Caption         =   "3."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   960
      Width           =   135
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00000000&
      Caption         =   "4."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00000000&
      Caption         =   "5."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label lblSeconds1 
      BackColor       =   &H00000000&
      Caption         =   "seconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblSeconds2 
      BackColor       =   &H00000000&
      Caption         =   "seconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblSeconds3 
      BackColor       =   &H00000000&
      Caption         =   "seconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblSeconds4 
      BackColor       =   &H00000000&
      Caption         =   "seconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblSeconds5 
      BackColor       =   &H00000000&
      Caption         =   "seconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "frmHardHS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This takes you back
Private Sub cmdBack_Click()
    frmHardHS.Hide
End Sub
'This checks to see if you've made the high scores
Private Sub cmdCheck_Click()
    If HardTime < Val(txt1.Text) Then
        txt5.Text = txt4.Text
        txt4.Text = txt3.Text
        txt3.Text = txt2.Text
        txt2.Text = txt1.Text
        txt1.Text = HardTime
        
        txtname5.Text = txtname4.Text
        txtname4.Text = txtname3.Text
        txtname3.Text = txtname2.Text
        txtname2.Text = txtname1.Text
        txtname1.Text = HardName
    ElseIf HardTime < Val(txt2.Text) Then
        txt5.Text = txt4.Text
        txt4.Text = txt3.Text
        txt3.Text = txt2.Text
        txt2.Text = HardTime
        
        txtname5.Text = txtname4.Text
        txtname4.Text = txtname3.Text
        txtname3.Text = txtname2.Text
        txtname2.Text = HardName
    ElseIf HardTime < Val(txt3.Text) Then
        txt5.Text = txt4.Text
        txt4.Text = txt3.Text
        txt3.Text = HardTime
        
        txtname5.Text = txtname4.Text
        txtname4.Text = txtname3.Text
        txtname3.Text = HardName
    ElseIf HardTime < Val(txt4.Text) Then
        txt5.Text = txt4.Text
        txt4.Text = HardTime
        
        txtname5.Text = txtname4.Text
        txtname4.Text = HardName
    ElseIf HardTime < Val(txt5.Text) Then
        txt5.Text = HardTime
        
        txtname5.Text = HardName
    Else
        MsgBox "Sorry" & FirstName & ", your time did not make the high scores for the easy puzzle", , "No High Score"
    End If
End Sub
