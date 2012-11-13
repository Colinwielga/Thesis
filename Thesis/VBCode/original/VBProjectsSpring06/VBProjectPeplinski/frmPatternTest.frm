VERSION 5.00
Begin VB.Form frmPatternTest 
   BackColor       =   &H0080FF80&
   Caption         =   "Pattern Test"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   ForeColor       =   &H0080FF80&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Opt9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   12
      Top             =   6960
      Width           =   735
   End
   Begin VB.OptionButton Opt8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   6960
      Width           =   855
   End
   Begin VB.OptionButton Opt7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   6960
      Width           =   735
   End
   Begin VB.OptionButton Opt6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   4560
      Width           =   735
   End
   Begin VB.OptionButton Opt5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   4560
      Width           =   735
   End
   Begin VB.OptionButton Opt4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.OptionButton Opt3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return To Main Menu"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdCompute 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label lblImages 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Images displayed by Pagina Blu"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Label lblChoice 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Choose the image that most appeals to you"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   7680
      Width           =   3975
   End
   Begin VB.Image Image9 
      Height          =   1695
      Left            =   5400
      Picture         =   "frmPatternTest.frx":0000
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Image Image8 
      Height          =   1695
      Left            =   3120
      Picture         =   "frmPatternTest.frx":9656
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Image Image7 
      Height          =   1695
      Left            =   720
      Picture         =   "frmPatternTest.frx":12CAC
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   1695
      Left            =   5520
      Picture         =   "frmPatternTest.frx":1C302
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   1695
      Left            =   3120
      Picture         =   "frmPatternTest.frx":25958
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   1695
      Left            =   720
      Picture         =   "frmPatternTest.frx":2EFAE
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1695
      Left            =   5400
      Picture         =   "frmPatternTest.frx":38604
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   3000
      Picture         =   "frmPatternTest.frx":41C5A
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   720
      Picture         =   "frmPatternTest.frx":4B2B0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmPatternTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
'Returns the user to the main menu
    frmPatternTest.Hide
    frmBegin.Show
End Sub

Private Sub cmdCompute_Click()
'this command button determines what output should go with the user's chosen image
'in describing which image most appeals to him or her
    
    'declare variables
    Dim Answer As Integer, Pos As Integer, Ctr As Integer
    Dim Size As Integer
    Dim Found As Boolean
    Dim NumArray(1 To 9) As Integer
    Dim PictureArray(1 To 9) As String
   
   'find which option button the user chose and this will be the search value for
   'the match and stop search within the array
    If Opt1.Value = True Then
        Answer = 1
    End If
    If Opt2.Value = True Then
        Answer = 2
    End If
    If Opt3.Value = True Then
        Answer = 3
    End If
    If Opt4.Value = True Then
        Answer = 4
    End If
    If Opt5.Value = True Then
        Answer = 5
    End If
    If Opt6.Value = True Then
        Answer = 6
    End If
    If Opt7.Value = True Then
        Answer = 7
    End If
    If Opt8.Value = True Then
        Answer = 8
    End If
    If Opt9.Value = True Then
        Answer = 9
    End If
    
    
    Pos = 0
    Size = 9
    Ctr = 0
    Found = False
    
    'Open file into an array
    Open App.Path & "\PatternResults.txt" For Input As #1
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, NumArray(Pos), PictureArray(Pos)
    Loop
    Close #1
    
    'match and stop search to determine what picture the user chose
    Do While (Found = False And Ctr < Size)
        Ctr = Ctr + 1
        If NumArray(Ctr) = Answer Then
            Found = True
        End If
    Loop
    
    'Display results of the Pattern Test
    If Found = True Then
        MsgBox PictureArray(Ctr), , "Results of Pattern Test"
    Else
        MsgBox "Entered an incorrect number", , "Error"
    End If
End Sub


