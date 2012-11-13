VERSION 5.00
Begin VB.Form frm3 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Grains"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdrefined 
      Caption         =   "Show Some Refined Grain Examples"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdwhole 
      Caption         =   "Show Some Whole Grain Examples"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back To Pyramid"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox picoutput 
      Height          =   4815
      Left            =   2880
      ScaleHeight     =   4755
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Depending on your age, it is recommended that you have six to eight servings of grains per day"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7680
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'frm3
'Ben Morris
'March 21
'displays the different grains and sorts them
Option Explicit
Dim whole(1 To 14) As String
Dim refined(1 To 14) As String
Dim CTR As Integer
Dim sort As Integer

Private Sub cmdback_Click()
    frm1.Show
    frm2.Hide
    frm3.Hide
    frm4.Hide
    frm5.Hide
    frm6.Hide
    frm7.Hide
    frm8.Hide
    'goes back to the pyramid and hides all other forms
End Sub

Private Sub cmdclear_Click()
    picoutput.Cls
End Sub

Private Sub cmdrefined_Click()
    picoutput.Cls
    CTR = 0
    picoutput.Print "These are some Examples of Refined Grains"
    picoutput.Print "--------------------------------------------------------------------"
    Open App.Path & "\RefinedGrains.txt" For Input As #1
    Do Until EOF(1)
    
    ' Get the data from the file
        CTR = CTR + 1
        Input #1, refined(CTR)
        
        picoutput.Print refined(CTR)
        
    'this is the code to bubble sort the grains
    Loop
    Close #1
    sort = InputBox("Would you like to sort alphabetically? Enter 1 for yes or 2 for no", "Sort")
        If sort = 1 Then
            Dim pass As Integer, pos As Integer, temp As String
    
    picoutput.Cls
    picoutput.Print "Alphabetically Sorted Refined Grains"
    picoutput.Print "------------------------------------------------------------------------------------------"
    
    For pass = 1 To CTR - 1
        For pos = 1 To CTR - pass
            If refined(pos) > refined(pos + 1) Then
                temp = refined(pos)
                refined(pos) = refined(pos + 1)
                refined(pos + 1) = temp
            End If
        Next pos
    Next pass
    
    For pass = 1 To CTR
        picoutput.Print refined(pass)
    Next pass
    
    Else
        MsgBox "OK", , "Done"
    End If
End Sub


Private Sub cmdwhole_Click()
    picoutput.Cls
    CTR = 0
    picoutput.Print "These are some Examples of Whole Grains"
    picoutput.Print "--------------------------------------------------------------------"
    Open App.Path & "\Grains.txt" For Input As #1
    Do Until EOF(1)
    
    ' Get the data from the file
        CTR = CTR + 1
        Input #1, whole(CTR)
        
        picoutput.Print whole(CTR)
    
    Loop
    Close #1
    
    'this also bubble sorts the grains
    sort = InputBox("Would you like to sort alphabetically? Enter 1 for yes or 2 for no", "Sort")
        If sort = 1 Then
            Dim pass As Integer, pos As Integer, temp As String
    
    picoutput.Cls
    picoutput.Print "Alphabetically Sorted Whole Grains"
    picoutput.Print "------------------------------------------------------------------------------------------"
    
    For pass = 1 To CTR - 1
        For pos = 1 To CTR - pass
            If whole(pos) > whole(pos + 1) Then
                temp = whole(pos)
                whole(pos) = whole(pos + 1)
                whole(pos + 1) = temp
            End If
        Next pos
    Next pass
    
    For pass = 1 To CTR
        picoutput.Print whole(pass)
    Next pass
    
    Else
        MsgBox "OK", , "Done"
    End If
    
End Sub

