VERSION 5.00
Begin VB.Form frmFindGames 
   Caption         =   "Look Up Games"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "OCR A Extended"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSortPS3 
      Caption         =   "Sort"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSortWii 
      Caption         =   "Sort"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   2400
      ScaleHeight     =   6195
      ScaleWidth      =   9315
      TabIndex        =   5
      Top             =   120
      Width           =   9375
   End
   Begin VB.CommandButton cmdSort360 
      Caption         =   "Sort"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturnToWant 
      Caption         =   "Return"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdPS3 
      Caption         =   "PS3"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdWii 
      Caption         =   "Wii"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmdXbox360 
      Caption         =   "Xbox 360"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmFindGames.frx":0000
      Top             =   -360
      Width           =   12000
   End
End
Attribute VB_Name = "frmFindGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmFindGames
'26 March 2007

'This form allows users to see New and Current games for Xbox 360, Playstation 3 and Nintendo Wii.
'Clicking the respected console's command buttun, a list of games will be displayed in a picture box.
'Users may also choose to sort games alphabetically.

Option Explicit
'This command opens the PS3.txt file and displays
' the game titles featured in the picture box.
Private Sub cmdPS3_Click()
Dim Ctr As Integer
    Open App.Path & "\PS3.txt" For Input As #1      'Opens txt document for display
        picResults.Cls
        Ctr = 0
        picResults.Print "Title", , , "Platform", , , "Release"
        picResults.Print "*****************************************************************************************************************************"
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, Title(Ctr), Platform(Ctr), Release(Ctr)
        picResults.Print Title(Ctr); Tab(42); Platform(Ctr); Tab(85); Release(Ctr)
    Loop
    picResults.Print "*********************************************************************************************************************************"
    Close #1
End Sub
Private Sub cmdReturnToWant_Click()
    frmFindGames.Hide       'Hides FindGames form
    frmSelectWant.Show      'Shows SelectWant form
End Sub
'This Command allows users to sort Xbox 360 titles alphabetically
Private Sub cmdSort360_Click()
    Dim Ctr As Integer      'Dimmed variables
    Dim Pass As Integer
    Dim Pos As Integer
    Dim Temp As String
    Dim Title(1 To 100) As String
    Open App.Path & "\Xbox3602.txt" For Input As #1     'Opens txt document for display
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Title(Ctr), Platform(Ctr), Release(Ctr)
            picResults.Print Title(Ctr); Tab(42); Platform(Ctr); Tab(85); Release(Ctr)
        Loop
            picResults.Print "Title", , , "Platform", , , "Release"
            picResults.Print "*****************************************************************************************************************************"
        Close #1
        picResults.Cls
        For Pass = 1 To (Ctr - 1)
            For Pos = 1 To (Ctr - Pass)
                If Title(Pos) > Title(Pos + 1) Then
                    Temp = Title(Pos)
                    Title(Pos) = Title(Pos + 1)
                    Title(Pos + 1) = Temp
                End If
            Next Pos
        Next Pass
    For Pos = 1 To Ctr
        picResults.Print Title(Pos)
    Next Pos
End Sub

Private Sub cmdSortPS3_Click()
    Dim Ctr As Integer
    Dim Pass As Integer
    Dim Pos As Integer
    Dim Temp As String
    Dim Title(1 To 100) As String
    Open App.Path & "\PS3.txt" For Input As #1      'Opens txt document for display
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Title(Ctr), Platform(Ctr), Release(Ctr)
            picResults.Print Title(Ctr); Tab(42); Platform(Ctr); Tab(85); Release(Ctr)
        Loop
            picResults.Print "Title", , , "Platform", , , "Release"
            picResults.Print "*****************************************************************************************************************************"
        Close #1
        picResults.Cls
        For Pass = 1 To (Ctr - 1)
            For Pos = 1 To (Ctr - Pass)
                If Title(Pos) > Title(Pos + 1) Then
                    Temp = Title(Pos)
                    Title(Pos) = Title(Pos + 1)
                    Title(Pos + 1) = Temp
                End If
            Next Pos
        Next Pass
    For Pos = 1 To Ctr
        picResults.Print Title(Pos)
    Next Pos
End Sub

Private Sub cmdSortWii_Click()
    Dim Ctr As Integer
    Dim Pass As Integer
    Dim Pos As Integer
    Dim Temp As String
    Dim Title(1 To 100) As String
    Open App.Path & "\Wii.txt" For Input As #1      'Opens txt document for display
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Title(Ctr), Platform(Ctr), Release(Ctr)
            picResults.Print Title(Ctr); Tab(42); Platform(Ctr); Tab(85); Release(Ctr)
        Loop
            picResults.Print "Title", , , "Platform", , , "Release"
            picResults.Print "*****************************************************************************************************************************"
        Close #1
        picResults.Cls
        For Pass = 1 To (Ctr - 1)
            For Pos = 1 To (Ctr - Pass)
                If Title(Pos) > Title(Pos + 1) Then
                    Temp = Title(Pos)
                    Title(Pos) = Title(Pos + 1)
                    Title(Pos + 1) = Temp
                End If
            Next Pos
        Next Pass
    For Pos = 1 To Ctr
        picResults.Print Title(Pos)
    Next Pos
End Sub
'This command opens the Wii.txt file and displays
'the game titles featured in the picture box.
Private Sub cmdWii_Click()
Dim Ctr As Integer
    Open App.Path & "\Wii.txt" For Input As #1      'Opens txt document for display
        picResults.Cls
        Ctr = 0
        picResults.Print "Title", , , "Platform", , , "Release"
        picResults.Print "*****************************************************************************************************************************"
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, Title(Ctr), Platform(Ctr), Release(Ctr)
        picResults.Print Title(Ctr); Tab(42); Platform(Ctr); Tab(85); Release(Ctr)
    Loop
    picResults.Print "*********************************************************************************************************************************"
    Close #1
End Sub
'This command opens the Xbox3602.txt file and displays
'the game system featured in the picture box.
Private Sub cmdXbox360_Click()
Dim Ctr As Integer
    Open App.Path & "\Xbox3602.txt" For Input As #1     'Opens txt document for display
        picResults.Cls
        Ctr = 0
        picResults.Print "Title", , , "Platform", , , "Release"
        picResults.Print "*****************************************************************************************************************************"
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, Title(Ctr), Platform(Ctr), Release(Ctr)
        picResults.Print Title(Ctr); Tab(42); Platform(Ctr); Tab(85); Release(Ctr)
    Loop
    picResults.Print "*********************************************************************************************************************************"
    Close #1
End Sub

