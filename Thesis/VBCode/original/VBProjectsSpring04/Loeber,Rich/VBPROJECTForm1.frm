VERSION 5.00
Begin VB.Form frmMusicEditor 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   12
      Text            =   "3. Click on the Alphabetize Songs to Alphabetize the file names.  Then on ""After Program"" to see a final result of the file names."
      Top             =   3120
      Width           =   10215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Text            =   "2.  Click on the ""Prior To Program"" button to see what your files will look like before formatting."
      Top             =   2640
      Width           =   7695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "After Program"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   7
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Alphabetize Songs"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Prior To Program"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Text            =   "1.  Click on the ""Run"" button which will format the files and delete unnecessary characters."
      Top             =   2160
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox picResults20 
      BackColor       =   &H80000013&
      Height          =   2055
      Left            =   2280
      ScaleHeight     =   1995
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   6000
      Width           =   7935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000013&
      Height          =   1935
      Left            =   2280
      ScaleHeight     =   1875
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   3840
      Width           =   7935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "MUSIC EDITOR"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Author: Rich Loeber"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "Directions:"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
   End
End
Attribute VB_Name = "frmMusicEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name is "Music Editor"
'Form Name (frmMusicEditor)
'Author "Rich Loeber"
'Due Date March 15, 2004
'Purpose: When downloading music I have found that the file names are normally altered and there are
          'unwanted characters that cause the files not be in alphabetical order.
          'This program will allow you to place your music into a certain folder and get rid of
          'the unwanted characters and alphabetize them.

Option Explicit
Public Path As String
Dim Path1 As String

Private Sub cmdquit_Click()
End
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Run" Then
    Rename
Else
    Command1.Caption = "Run"
End If
End Sub


Private Sub Rename()
Dim FileName As String
Dim PreviousName As String
Dim Path As String
Dim Path2 As String
Dim I As Integer
Dim CTR As Integer
Dim Junk(1 To 100) As String

'This button should read the information in the folder and format the names into the proper FileName
'The only files available to me on the network came from Windows Media Player.  If the format of the files will not work
'I can send some music files to you.

Path = "N:\CS130\handin\Loeber,Rich\Music and MP3s\"
Path2 = "N:\CS130\handin\Loeber, Rich\"
FileName = (Path & ".wmz")
CTR = 0

Open (Path2 & "junk.txt") For Input As #1
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, Junk(CTR)
       Loop
End Sub


Private Sub Command2_Click()
Path1 = "N:\CS130\handin\Loeber, Rich\Music and MP3s\"
picResults.Cls
picResults.Picture = LoadPicture(Path1 & "PriorToProgram.JPG")
End Sub

Private Sub Command3_Click()
Path1 = "N:\CS130\handin\Loeber, Rich\Music and MP3s\"
picResults.Cls
picResults20.Picture = LoadPicture(Path1 & "AfterProgram.JPG")
End Sub

Private Sub Command4_Click()
Dim Pass As Integer
Dim FileName(1 To 100) As String
Dim Temp As String
Dim Compare As Integer
picResults.Cls
For Pass = 1 To 100
    For Compare = 1 To 100 - Pass
        If FileName(Compare) > FileName(Compare + 1) Then
        Temp = FileName(Compare)
        FileName(Compare) = FileName(Compare + 1)
        FileName(Compare + 1) = Temp
    End If
Next Compare
Next Pass
End Sub


