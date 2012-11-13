VERSION 5.00
Begin VB.Form frm2 
   BackColor       =   &H00000000&
   Caption         =   "Main Page"
   ClientHeight    =   10605
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   10605
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMinneapolis 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   975
      Left            =   3600
      TabIndex        =   5
      Text            =   "Minneapolis Travel Guide"
      Top             =   240
      Width           =   8535
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FF8080&
      Caption         =   "Search"
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      Height          =   855
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmdGame 
      BackColor       =   &H00FF8080&
      Caption         =   "Test your Knowledge"
      Height          =   855
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmdMap 
      BackColor       =   &H00FF8080&
      Caption         =   "View Map"
      Height          =   855
      Left            =   2880
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9000
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   1680
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   7035
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   1440
      Width           =   12015
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   855
      Left            =   4560
      SourceDoc       =   "N:\Classes\MoveStuff\Kotila_Rowe\StuartDavis1"
      TabIndex        =   6
      Top             =   9000
      Width           =   1695
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minneapolis Travel Guide
'frm1- Map
'Kayla Kotila and Chris Rowe
'March 30
'This is main page, has buttons to go to other forms and search function

'
Option Explicit
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdGame_Click()
'change forms
Form1.Show
frm2.Hide

End Sub

Private Sub cmdMap_Click()
'change forms
frm2.Hide
frm1.Show

End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSearch_Click()
   'Match-stop search w/ input array
    Dim Places(1 To 26) As String
    Dim CTR As Integer
    Dim Place As String
    Dim Found As Boolean
    Dim N As Integer
    N = 1
    CTR = 0
    Found = False
    
    Place = InputBox("Enter a Place", "Search")
    

    Open App.Path & "\List.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Places(CTR)
    Loop
    Close #1
    
        
     Do Until (Found = True And N > CTR)
        If Places(N) = Place Then
        Found = True
        End If
        N = N + 1
    Loop
   
    If Found Then
        MsgBox "We included that in our list!", vbExclamation, "Search"
        N = 1
    Else
        MsgBox "Sorry, We did not include that place in out list.", vbCritical, "Search"
        N = 1
    End If
    
End Sub



