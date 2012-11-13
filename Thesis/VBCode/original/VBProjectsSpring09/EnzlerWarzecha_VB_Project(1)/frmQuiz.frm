VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ck6 
      BackColor       =   &H0080C0FF&
      Caption         =   "You practice confidentiality"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   8
      Top             =   3840
      Width           =   3495
   End
   Begin VB.CheckBox ck5 
      BackColor       =   &H0080C0FF&
      Caption         =   "You like doughnuts"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   7
      Top             =   5040
      Width           =   3495
   End
   Begin VB.CheckBox ck4 
      BackColor       =   &H0080C0FF&
      Caption         =   "You would like wearing a uniform to work"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   6
      Top             =   6720
      Width           =   3495
   End
   Begin VB.CheckBox ck3 
      BackColor       =   &H0080C0FF&
      Caption         =   "You speed when you drive"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CheckBox ck2 
      BackColor       =   &H0080C0FF&
      Caption         =   "You like enforcing the rules"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CheckBox ck1 
      BackColor       =   &H0080C0FF&
      Caption         =   "You've been arrested"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   3
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CommandButton cmdScore 
      BackColor       =   &H0080C0FF&
      Caption         =   "Score The Test!"
      Height          =   2055
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   1320
      Picture         =   "frmQuiz.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Please check all that apply to you:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   9
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Are You Cut Out For A Career In Law Enforcement?"
      BeginProperty Font 
         Name            =   "Myriad Web Pro Condensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5520
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu return 
         Caption         =   "Return to Menu"
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form will ask the user to check multiple check boxes
'If the user checks yes, a "1" will be added to the Total
'The total will determine whether the user is cut out for law enforcement


Option Explicit
Dim Total As Integer

'Crime Awareness Project
'frm Quiz
'By: Alex Warzecha & Andrew Enzler
'Written: 3/18/2009

Private Sub cmdScore_Click()

Total = 0 'Set total as zero as no questions have been answered yet

If ck1.Value = "1" Then 'A 1 isn't added in this line because this is not a desirable trait
Else: ck1.Value = "2"
End If

If ck2.Value = "1" Then
    Total = Total + 1 'If the check box is marked then 1 is added to the total
Else: ck2.Value = "2"
End If

If ck3.Value = "1" Then
Else: ck3.Value = "2"
End If

If ck4.Value = "1" Then
    Total = Total + 1
Else: ck4.Value = "2"
End If

If ck5.Value = "1" Then
    Total = Total + 1
Else: ck5.Value = "2"
End If

If ck6.Value = "1" Then
    Total = Total + 1
Else: ck6.Value = "2"
End If



If Total >= 4 Then ' if user got 4 right, he/she will get this message
'This form decides by the number the user 'got right' whether he/she should be a police officer
    
    MsgBox "Law Enforcement is the perfect career for you", , "Congratulations!"
ElseIf Total >= 2 Then
    MsgBox "Law Enforcement may or may not be the career for you", , "Questionable"
ElseIf Total < 2 Then
    MsgBox "Law Enforcement is not the career for you", , "No Way"
End If

End Sub
Private Sub quit_Click()
End 'Ends Program
End Sub

Private Sub return_Click()
frmQuiz.Hide
frmHome.Show 'Will return to home form
End Sub


