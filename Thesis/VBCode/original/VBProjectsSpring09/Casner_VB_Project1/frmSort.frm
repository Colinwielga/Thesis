VERSION 5.00
Begin VB.Form frmSort 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBoth 
      Caption         =   "Sort both data sets"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSecondary 
      Caption         =   "Sort secondary set"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrimary 
      Caption         =   "Sort primary set"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblSort 
      BackColor       =   &H00008000&
      Caption         =   "Which data set do you wish to sort?"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Visual Basic T-test
'Benjamin Casner
'March 13th, 2009
'frmSort
'this form allows the user to decide which data set to sort
'The subroutines (except for the form load, which makes sure that
'this form is not visible as the program first runs) are
'basic sorting and swap algorithms
Private Sub cmdBoth_Click()
    'sorts both sets
    Dim temp As Single, pos As Integer, pass As Integer
    frmSort.Hide
    For pass = 1 To ctr1 - 1
        For pos = 1 To ctr1 - pass
            If Sample1(pos) > Sample1(pos + 1) Then
                temp = Sample1(pos + 1)
                Sample1(pos + 1) = Sample1(pos)
                Sample1(pos) = temp
            End If
        Next pos
    Next pass
    For pass = 1 To ctr2 - 1
        For pos = 1 To ctr2 - pass
            If Sample2(pos) > Sample2(pos + 1) Then
                temp = Sample2(pos + 1)
                Sample2(pos + 1) = Sample2(pos)
                Sample2(pos) = temp
            End If
        Next pos
    Next pass
End Sub

Private Sub cmdPrimary_Click()
    'sorts primary set
    Dim temp As Single, pos As Integer, pass As Integer
    frmSort.Hide
    For pass = 1 To ctr1 - 1
        For pos = 1 To ctr1 - pass
            If Sample1(pos) > Sample1(pos + 1) Then
                temp = Sample1(pos + 1)
                Sample1(pos + 1) = Sample1(pos)
                Sample1(pos) = temp
            End If
        Next pos
    Next pass
End Sub

Private Sub cmdSecondary_Click()
    'sort secondary set
    Dim temp As Single, pos As Integer, pass As Integer
    frmSort.Hide
    For pass = 1 To ctr2 - 1
        For pos = 1 To ctr2 - pass
            If Sample2(pos) > Sample2(pos + 1) Then
                temp = Sample2(pos + 1)
                Sample2(pos + 1) = Sample2(pos)
                Sample2(pos) = temp
            End If
        Next pos
    Next pass
End Sub

Private Sub Form_Load()
    frmDataDisplay.Hide
End Sub
