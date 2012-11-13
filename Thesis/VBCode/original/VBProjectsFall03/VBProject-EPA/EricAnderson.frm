VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF00&
   Caption         =   "USS Time Standards"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   7920
      Picture         =   "Eric Anderson.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00FFFF80&
      Height          =   3255
      Left            =   4440
      ScaleHeight     =   3195
      ScaleWidth      =   3555
      TabIndex        =   5
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   8880
      TabIndex        =   4
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort Swimmers By Time (Fastest to Slowest)"
      Height          =   1095
      Left            =   840
      TabIndex        =   3
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Out Time Standards Achieved"
      Height          =   1215
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Data File Of Times and Names"
      Height          =   1215
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Eric Anderson"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "USS Swimming Standard Times"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'USS Time Standards, Eric Anderson.frm'
'Eric Anderson'
'October 29, 2003'
'The purpose is to find out which time standards different swimmers fit into, and then sort them by how fast they are.'
Option Explicit
Dim strname(1 To 10) As String
Dim time(1 To 10) As Single
Public strpath As String

Private Sub cmdPrint_Click()
Dim i As Integer
pbxResults.Cls
For i = 1 To 10 'read each array'
    If time(i) > 29.19 Then
        pbxResults.Print strname(i); " achieved a C time."
    ElseIf time(i) > 27.09 Then
        pbxResults.Print strname(i); " achieved a B time."
    ElseIf time(i) > 24.99 Then
        pbxResults.Print strname(i); " achieved a BB time."
    ElseIf time(i) > 23.99 Then
        pbxResults.Print strname(i); " achieved an A time."
    ElseIf time(i) > 22.89 Then
        pbxResults.Print strname(i); " achieved an AA time."
    ElseIf time(i) > 21.89 Then
        pbxResults.Print strname(i); " achieved an AAA time."
    ElseIf time(i) > 21.09 Then
        pbxResults.Print strname(i); " achieved an AAAA time."
    End If
Next i
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()
Dim i As Integer
    Open strpath & "SwimmingTimes.txt" For Input As #1 'open file to read'
        For i = 1 To 10
            Input #1, strname(i), time(i) 'read each array'
        Next i
    Close #1
End Sub

Private Sub cmdSort_Click()
Dim pass As Integer
Dim temp1 As Single
Dim temp2 As String
Dim N As Integer
Dim i As Integer
N = 10
pbxResults.Cls
For pass = 1 To (N - 1) 'sort by time'
    For i = 1 To (N - pass)
        If time(i) > time(i + 1) Then
            temp1 = time(i)
            time(i) = time(i + 1)
            time(i + 1) = temp1
            temp2 = strname(i)
            strname(i) = strname(i + 1)
            strname(i + 1) = temp2
        End If
    Next i
Next pass
For i = 1 To 10 'print each person and their time from fastest to slowest'
    pbxResults.Print strname(i), Tab(30); FormatNumber(time(i), 2); " sec."
Next i
        
End Sub

Private Sub Form_Load()
strpath = "N:\CS130\handin\VB Project\" 'strpath makes file easily accessible'
End Sub
