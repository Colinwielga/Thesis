VERSION 5.00
Begin VB.Form frmGameInformation 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Game Information"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearchPS3 
      Caption         =   "Search PS3 Titles"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton cmdSearchWii 
      Caption         =   "Search Wii Titles"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
   End
   Begin VB.PictureBox picResults 
      Height          =   5295
      Left            =   3720
      ScaleHeight     =   5235
      ScaleWidth      =   7875
      TabIndex        =   3
      Top             =   3960
      Width           =   7935
   End
   Begin VB.CommandButton cmdSearch360 
      Caption         =   "Search 360 Titles"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label lblGameInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Looking for a particular game or information? Find it here!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   13695
   End
   Begin VB.Image Image1 
      Height          =   4815
      Left            =   0
      Picture         =   "frmGameInformation.frx":0000
      Top             =   0
      Width           =   8250
   End
End
Attribute VB_Name = "frmGameInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmMain
'26 March 2007

'This form allows users to search for a particular game and may do so according to system.
Option Explicit
Private Sub cmdReturn_Click()
    frmGameInformation.Hide     'Hides GameInformation form
    frmSelectWant.Show          'Shows SelectWant form
End Sub
Private Sub cmdSearch360_Click()
    Dim Pass As Integer
    Dim Ctr As Integer
    Dim Pos As Integer
    Dim Temp As String
    Dim Title(1 To 100) As String
    Dim Game As Integer
    
    Open App.Path & "\Xbox3602.txt" For Input As #1
    
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Title(Ctr), Platform(Ctr), Release(Ctr)
    Loop
    For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If Title(Pos) > Title(Pos + 1) Then
                Temp = Title(Pos)
                Title(Pos) = Title(Pos + 1)
                Title(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    For Game = 1 To Ctr
        picResults.Print Title(Game)
    Next Game
End Sub

Private Sub cmdSearchWii_Click()
    Dim Pass As Integer
    Dim Ctr As Integer
    Dim Pos As Integer
    Dim Temp As String
    Dim Title(1 To 100) As String
    Dim Game As Integer
    
    Open App.Path & "\Wii.txt" For Input As #1
    
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Title(Ctr), Platform(Ctr), Release(Ctr)
    Loop
    For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If Title(Pos) > Title(Pos + 1) Then
                Temp = Title(Pos)
                Title(Pos) = Title(Pos + 1)
                Title(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    For Game = 1 To Ctr
        picResults.Print Title(Game)
    Next Game
End Sub
