VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   Caption         =   "Form2"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form2"
   ScaleHeight     =   5100
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgondi 
      Caption         =   "Quit"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdtothebegin 
      Caption         =   "Go back to the beginning."
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to find your site!"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Text            =   "www.google.com"
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Building a Cedar Strip Canoe: www.aracnet.com/~ncglad/canoeentry.htm "
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   5295
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "Carrying Place Canoe Works: www.carryingplacecanoeworks.on.ca"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Caption         =   "The Newfound Woodworks: www.newfound.com"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Wooden Canoe Heritage Association: www.wcha.org"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label lblLink 
      Caption         =   "lblLink"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Minnesota Canoe Association: www.canoe-kayak.org/"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   $"Linkform.frx":0000
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Enter &URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ****************************************************************
'  Copyright ©1997-99, Karl E. Peterson
'  http://www.mvps.org/vb
' ****************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' ****************************************************************
Option Explicit
'Purpose = This form is here so that the user can begin their research on building thier own canoe.
'Project Name= Visual Basic Canoe Project("M:\CS130\CanoeProject")
'Form Name= Form 1 ("M:\CS130\CanoeProject\Project1.vbp\Linkform.frm")
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function HyperJump(ByVal URL As String) As Long
   HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Private Sub cmdgondi_Click()
End
End Sub
Private Sub cmdtothebegin_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Command1_Click()
   Call HyperJump(Text1.Text)
End Sub

Private Sub Form_Load()
   Text1.Text = "http://www.google.com"
   Me.Icon = Nothing
   
   ' make link label look like a link
   With lblLink
      .Font.Underline = True
      .ForeColor = vbBlue
      .MousePointer = 99 'custom
      Set .MouseIcon = Command1.MouseIcon
   End With
End Sub

Private Sub lblLink_Click()
   Call HyperJump(lblLink.Caption)
End Sub

Private Sub Text1_Change()
   lblLink.Caption = Text1.Text
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call HyperJump(Text1.Text)
      KeyAscii = 0
   End If
End Sub
