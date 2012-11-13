VERSION 5.00
Begin VB.Form BigFourFirms 
   BackColor       =   &H00008080&
   Caption         =   "Form3"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form3"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   7800
      ScaleHeight     =   915
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton cmdswitch 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ready? Let see now how much you know about accounting!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   2655
   End
   Begin VB.CommandButton CmdArrange 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Who are the big four?"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   2775
   End
   Begin VB.PictureBox picResults4 
      Height          =   9615
      Left            =   7560
      ScaleHeight     =   9555
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   1800
      Width           =   6135
      Begin VB.PictureBox Picture4 
         Height          =   1935
         Left            =   3600
         ScaleHeight     =   1875
         ScaleWidth      =   2235
         TabIndex        =   8
         Top             =   4320
         Width           =   2295
      End
      Begin VB.PictureBox Picture3 
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2595
         ScaleWidth      =   3315
         TabIndex        =   7
         Top             =   4200
         Width           =   3375
      End
      Begin VB.PictureBox Picture2 
         Height          =   1215
         Left            =   3000
         ScaleHeight     =   1155
         ScaleWidth      =   1515
         TabIndex        =   6
         Top             =   2760
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdBigFour 
      BackColor       =   &H00C0E0FF&
      Caption         =   "What are the Big Four?"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lblBigFour 
      BackColor       =   &H00FFC0FF&
      Caption         =   "          Big Four"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "BigFourFirms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Accounting basics and Income statement
'Form 4:The big Four
'Author:Patrick Niyibizi and Frankie Chan
'Date Written:October 8th 2009
'Objective:To provide information about the big four accountancy and auditing firms
Option Explicit
Private Sub CmdArrange_Click()
    Dim BigFour(1 To 5) As String, Revenues(1 To 5) As Single, Employees(1 To 5) As Double, CTR As Integer, L As Integer, pass As Integer, tempBigFour As String, tempRevenues As Single, tempEmployees As Double, pos As Integer   'Declare arrays and other variables used in this form
    CTR = 0
    Open App.Path & "\BigFour.txt" For Input As #1     'Open chanel and fill arrays
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, BigFour(CTR), Revenues(CTR), Employees(CTR)
        
    Loop
 Close #1       'Close chanel
 
For pass = 1 To CTR - 1                              'Sort elements in Arrays starting with highest revenue
    For pos = 1 To CTR - pass
        If Revenues(pos) < Revenues(pos + 1) Then
            tempRevenues = Revenues(pos)
            Revenues(pos) = Revenues(pos + 1)
            Revenues(pos + 1) = tempRevenues
            tempBigFour = BigFour(pos)
            BigFour(pos) = BigFour(pos + 1)
            BigFour(pos + 1) = tempBigFour
            tempEmployees = Employees(pos)
            Employees(pos) = Employees(pos + 1)
            Employees(pos + 1) = tempEmployees
         End If
    Next pos
Next pass

picResults4.Print "The Big Four in order of highest revenues(in billions) and their respective employees"          'Print title
picResults4.Print "----------------------------------------------------------------------------------------------------------------------------"
For L = 1 To CTR
    picResults4.Print BigFour(L); Tab(30); FormatCurrency(Revenues(L)); Tab(40); Employees(L)      'Print sorted arrays
Next

Picture1.Picture = LoadPicture(App.Path & "\Images\deloitte_logo.jpg")       'Print images of the logos of the big four firms
Picture2.Picture = LoadPicture(App.Path & "\Images\pwc_logo.gif")
Picture3.Picture = LoadPicture(App.Path & "\Images\KPMG_logo.jpg")
Picture4.Picture = LoadPicture(App.Path & "\Images\EY_logo.gif")
End Sub

Private Sub cmdBigFour_Click()       'Provide information about the big four using a message box
    MsgBox " the big four are the four largest international accountancy and professional services firms in the world."
End Sub

Private Sub cmdswitch_Click()             'Go to next form
BigFourFirms.Hide
Quizfrm.Show
End Sub
