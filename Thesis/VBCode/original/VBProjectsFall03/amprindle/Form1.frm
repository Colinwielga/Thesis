VERSION 5.00
Begin VB.Form FRMAngiePrindle 
   BackColor       =   &H80000002&
   Caption         =   "Angie's Family"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PBXResults 
      Height          =   4455
      Left            =   6480
      ScaleHeight     =   4395
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton CMDRead 
      BackColor       =   &H00C000C0&
      Caption         =   "Push me first!"
      Height          =   975
      Left            =   4800
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton CMDClear 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clear"
      Height          =   1095
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton CMDEnd 
      BackColor       =   &H0080C0FF&
      Caption         =   "End"
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton CMDAlph 
      BackColor       =   &H008080FF&
      Caption         =   "Sort Alphabetically"
      Height          =   1095
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton CMDEnter 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter an ID #"
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton CMDAge 
      BackColor       =   &H0000C000&
      Caption         =   "Sort by age"
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.PictureBox PBXPic 
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FRMAngiePrindle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:
        'To sort data by alphabetical order
        'To sort data by certain numerical order (age)
        'to find a picture of a certain person based on entry by user

        
Dim i As Integer
Dim StrNames(1 To 17) As String
Dim StrAges(1 To 17) As Integer
Dim StrNamestemp As String
Dim StrAgestemp As Integer
Public path As String
Dim NotFound As Boolean
Dim M As String
Dim StrPic(1 To 17) As String
Dim StrPictemp As String
Dim StrNum(1 To 17) As Integer
Dim StrNumtemp As Integer
Dim StrComm(1 To 17) As String
Dim StrCommtemp As String


Private Sub CMDAge_Click()
PBXResults.Cls
PBXResults.Print Tab(2); "ID"; Tab(15); "NAME", Tab(45); "AGE"
N = 17
For Pass = 1 To (N - 1)
For i = 1 To N - Pass
    If StrAges(i) > StrAges(i + 1) Then
        StrAgestemp = StrAges(i + 1)
        StrAges(i + 1) = StrAges(i)
        StrAges(i) = StrAgestemp

        StrPictemp = StrPic(i + 1)
        StrPic(i + 1) = StrPic(i)
        StrPic(i) = StrPictemp
        
        StrNamestemp = StrNames(i + 1)
        StrNames(i + 1) = StrNames(i)
        StrNames(i) = StrNamestemp
        
        StrNumtemp = StrNum(i + 1)
        StrNum(i + 1) = StrNum(i)
        StrNum(i) = StrNumtemp
        
        StrCommtemp = StrComm(i + 1)
        StrComm(i + 1) = StrComm(i)
        StrComm(i) = StrCommtemp
        
    End If
Next i
Next Pass
For i = 1 To 17

PBXResults.Print Tab(2); StrNum(i); ".)"; Tab(15); StrNames(i), Tab(45); StrAges(i)
Next i

End Sub

Private Sub CMDAlph_Click()
PBXResults.Cls
PBXResults.Print Tab(2); "ID"; Tab(15); "NAME", Tab(45); "AGE"
N = 17
For Pass = 1 To (N - 1)
For i = 1 To N - Pass
    If StrNames(i) > StrNames(i + 1) Then
        StrNamestemp = StrNames(i + 1)
        StrNames(i + 1) = StrNames(i)
        StrNames(i) = StrNamestemp

        StrAgestemp = StrAges(i + 1)
        StrAges(i + 1) = StrAges(i)
        StrAges(i) = StrAgestemp
        
        StrPictemp = StrPic(i + 1)
        StrPic(i + 1) = StrPic(i)
        StrPic(i) = StrPictemp
        
        StrNumtemp = StrNum(i + 1)
        StrNum(i + 1) = StrNum(i)
        StrNum(i) = StrNumtemp
        
        StrCommtemp = StrComm(i + 1)
        StrComm(i + 1) = StrComm(i)
        StrComm(i) = StrCommtemp
    End If
Next i
Next Pass

For i = 1 To 17
PBXResults.Print Tab(2); StrNum(i); ".)"; Tab(15); StrNames(i), Tab(45); StrAges(i)
Next i




End Sub

Private Sub CMDClear_Click()
PBXResults.Cls
End Sub

Private Sub CMDEnd_Click()
End
End Sub

Private Sub CMDEnter_Click()
Dim x As Double

M = InputBox("Enter a number")

i = 0
N = 17
NotFound = False
If M > 17 Then
    MsgBox "Sorry, number must be between 1 and 17", , "Error"
ElseIf M < 1 Then
    MsgBox "Sorry, number must be between 1 and 17", , "Error"
End If

Do While i <= N - 1 And NotFound = False
    i = i + 1
        If M = StrNum(i) Then
            NotFound = True
            x = i
        End If
    Loop

If NotFound = True Then
     PBXPic.Picture = LoadPicture((path & StrPic(i)))
     PBXResults.Cls
     PBXResults.Print StrComm(i)
End If



                                                           

End Sub

Private Sub CMDRead_Click()
Open path & "names2.txt" For Input As #1
For i = 1 To 17
    Input #1, StrNames(i), StrAges(i), StrPic(i), StrNum(i), StrComm(i)
    
Next i
Close #1

End Sub



Private Sub Form_Load()
path = "N:\CS130\handin\amprindle\"
End Sub
