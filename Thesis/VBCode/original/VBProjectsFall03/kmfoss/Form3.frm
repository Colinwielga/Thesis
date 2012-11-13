VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF8080&
   Caption         =   "Form3"
   ClientHeight    =   7695
   ClientLeft      =   510
   ClientTop       =   885
   ClientWidth     =   11895
   LinkTopic       =   "Form3"
   ScaleHeight     =   7695
   ScaleWidth      =   11895
   Begin VB.CommandButton cmdForm2 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   2055
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H00C0FFC0&
      Height          =   5895
      Left            =   2280
      ScaleHeight     =   5835
      ScaleWidth      =   9435
      TabIndex        =   7
      Top             =   1560
      Width           =   9495
   End
   Begin VB.CommandButton cmdSortHometown 
      Caption         =   "Sort by Hometown"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortBirthdate 
      Caption         =   "Sort By Birthdate"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortMajor 
      Caption         =   "Sort By Major"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search For a Resident"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortAlpha 
      Caption         =   "Sort Alphabetically"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read in Room Order"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   $"Form3.frx":0000
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RoomNumber(1 To 70) As Integer
Dim strName(1 To 70) As String
Dim Class(1 To 70) As String
Dim Major(1 To 70) As String
Dim Hometown(1 To 70) As String
Dim Birthdate(1 To 70) As String
Dim i As Integer
Dim Pass As Integer
Dim Temp As String
Dim N As Integer
Dim NotFound As Boolean
Dim Resident As String




Private Sub cmdClassStanding_Click()
Open strPath & "Roster.txt" For Input As #1
pbxResults.Cls
pbxResults.Print "Room Number"; Tab(25); "Name"; Tab(55); "Class"; Tab(65); "Major"; Tab(75); "Hometown"; Tab(100); "Birthdate"
pbxResults.Print "****************************************************************************************************************************************"
N = 41
    For Pass = 1 To N - 1
    For i = 1 To N - Pass
    
        If Class(i) > Class(i + 1) Then
        Temp = RoomNumber(i)
        RoomNumber(i) = RoomNumber(i + 1)
        RoomNumber(i + 1) = Temp
        Temp = strName(i)
        strName(i) = strName(i + 1)
        strName(i + 1) = Temp
        Temp = Class(i)
        Class(i) = Class(i + 1)
        Class(i + 1) = Temp
        Temp = Major(i)
        Major(i) = Major(i + 1)
        Major(i + 1) = Temp
        Temp = Hometown(i)
        Hometown(i) = Hometown(i + 1)
        Hometown(i + 1) = Temp
        Temp = Birthdate(i)
        Birthdate(i) = Birthdate(i + 1)
        Birthdate(i + 1) = Temp
        End If
    Next i
Next Pass

For i = 1 To N
pbxResults.Print RoomNumber(i); Tab(25); strName(i); Tab(55); Class(i); Tab(65); Major(i); Tab(75); Hometown(i); Tab(100); Birthdate(i)
Next i
Close #1
End Sub

Private Sub cmdForm2_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub cmdRead_Click()
pbxResults.Cls
pbxResults.Print "Room Number"; Tab(25); "Name"; Tab(55); "Class"; Tab(65); "Major"; Tab(75); "Hometown"; Tab(100); "Birthdate"
pbxResults.Print "****************************************************************************************************************************************"
Open strPath & "Roster.txt" For Input As #1
For i = 1 To 41
    Input #1, RoomNumber(i), strName(i), Class(i), Major(i), Hometown(i), Birthdate(i)
    pbxResults.Print RoomNumber(i); Tab(25); strName(i); Tab(55); Class(i); Tab(65); Major(i); Tab(75); Hometown(i); Tab(100); Birthdate(i)
    Next i
Close #1
End Sub

Private Sub cmdSearch_Click()
pbxResults.Cls
pbxResults.Print "Room Number"; Tab(25); "Name"; Tab(55); "Class"; Tab(65); "Major"; Tab(75); "Hometown"; Tab(100); "Birthdate"
pbxResults.Print "****************************************************************************************************************************************"
Resident = InputBox("Enter the name of the resident", "Name")
    i = 0
    NotFound = True
    Do While NotFound
        i = i + 1
        If Resident = strName(i) Then NotFound = False
    Loop
    If NotFound Then
        pbxResults.Print Resident; "was not found on this floor"
    Else
        pbxResults.Print RoomNumber(i); Tab(25); strName(i); Tab(55); Class(i); Tab(65); Major(i); Tab(75); Hometown(i); Tab(100); Birthdate(i)
    End If
End Sub

Private Sub cmdSortAlpha_Click()
Open strPath & "Roster.txt" For Input As #1
pbxResults.Cls
pbxResults.Print "Room Number"; Tab(25); "Name"; Tab(55); "Class"; Tab(65); "Major"; Tab(75); "Hometown"; Tab(100); "Birthdate"
pbxResults.Print "****************************************************************************************************************************************"
N = 41
    For Pass = 1 To N - 1
    For i = 1 To N - Pass
    
        If strName(i) > strName(i + 1) Then
        Temp = RoomNumber(i)
        RoomNumber(i) = RoomNumber(i + 1)
        RoomNumber(i + 1) = Temp
        Temp = strName(i)
        strName(i) = strName(i + 1)
        strName(i + 1) = Temp
        Temp = Class(i)
        Class(i) = Class(i + 1)
        Class(i + 1) = Temp
        Temp = Major(i)
        Major(i) = Major(i + 1)
        Major(i + 1) = Temp
        Temp = Hometown(i)
        Hometown(i) = Hometown(i + 1)
        Hometown(i + 1) = Temp
        Temp = Birthdate(i)
        Birthdate(i) = Birthdate(i + 1)
        Birthdate(i + 1) = Temp
        End If
    Next i
Next Pass

For i = 1 To N
pbxResults.Print RoomNumber(i); Tab(25); strName(i); Tab(55); Class(i); Tab(65); Major(i); Tab(75); Hometown(i); Tab(100); Birthdate(i)
Next i
Close #1
End Sub

Private Sub cmdSortBirthdate_Click()
Open strPath & "Roster.txt" For Input As #1
pbxResults.Cls
pbxResults.Print "Room Number"; Tab(25); "Name"; Tab(55); "Class"; Tab(65); "Major"; Tab(75); "Hometown"; Tab(100); "Birthdate"
pbxResults.Print "****************************************************************************************************************************************"
N = 41
    For Pass = 1 To N - 1
    For i = 1 To N - Pass
    
        If Birthdate(i) > Birthdate(i + 1) Then
        Temp = RoomNumber(i)
        RoomNumber(i) = RoomNumber(i + 1)
        RoomNumber(i + 1) = Temp
        Temp = strName(i)
        strName(i) = strName(i + 1)
        strName(i + 1) = Temp
        Temp = Class(i)
        Class(i) = Class(i + 1)
        Class(i + 1) = Temp
        Temp = Major(i)
        Major(i) = Major(i + 1)
        Major(i + 1) = Temp
        Temp = Hometown(i)
        Hometown(i) = Hometown(i + 1)
        Hometown(i + 1) = Temp
        Temp = Birthdate(i)
        Birthdate(i) = Birthdate(i + 1)
        Birthdate(i + 1) = Temp
        End If
    Next i
Next Pass

For i = 1 To N
pbxResults.Print RoomNumber(i); Tab(25); strName(i); Tab(55); Class(i); Tab(65); Major(i); Tab(75); Hometown(i); Tab(100); Birthdate(i)
Next i
Close #1
End Sub

Private Sub cmdSortHometown_Click()
Open strPath & "Roster.txt" For Input As #1
pbxResults.Cls
pbxResults.Print "Room Number"; Tab(25); "Name"; Tab(55); "Class"; Tab(65); "Major"; Tab(75); "Hometown"; Tab(100); "Birthdate"
pbxResults.Print "****************************************************************************************************************************************"
N = 41
    For Pass = 1 To N - 1
    For i = 1 To N - Pass
    
        If Hometown(i) > Hometown(i + 1) Then
        Temp = RoomNumber(i)
        RoomNumber(i) = RoomNumber(i + 1)
        RoomNumber(i + 1) = Temp
        Temp = strName(i)
        strName(i) = strName(i + 1)
        strName(i + 1) = Temp
        Temp = Class(i)
        Class(i) = Class(i + 1)
        Class(i + 1) = Temp
        Temp = Major(i)
        Major(i) = Major(i + 1)
        Major(i + 1) = Temp
        Temp = Hometown(i)
        Hometown(i) = Hometown(i + 1)
        Hometown(i + 1) = Temp
        Temp = Birthdate(i)
        Birthdate(i) = Birthdate(i + 1)
        Birthdate(i + 1) = Temp
        End If
    Next i
Next Pass

For i = 1 To N
pbxResults.Print RoomNumber(i); Tab(25); strName(i); Tab(55); Class(i); Tab(65); Major(i); Tab(75); Hometown(i); Tab(100); Birthdate(i)
Next i
Close #1
End Sub

Private Sub cmdSortMajor_Click()
Open strPath & "Roster.txt" For Input As #1
pbxResults.Cls
pbxResults.Print "Room Number"; Tab(25); "Name"; Tab(55); "Class"; Tab(65); "Major"; Tab(75); "Hometown"; Tab(100); "Birthdate"
pbxResults.Print "****************************************************************************************************************************************"
N = 41
    For Pass = 1 To N - 1
    For i = 1 To N - Pass
    
        If Major(i) > Major(i + 1) Then
        Temp = RoomNumber(i)
        RoomNumber(i) = RoomNumber(i + 1)
        RoomNumber(i + 1) = Temp
        Temp = strName(i)
        strName(i) = strName(i + 1)
        strName(i + 1) = Temp
        Temp = Class(i)
        Class(i) = Class(i + 1)
        Class(i + 1) = Temp
        Temp = Major(i)
        Major(i) = Major(i + 1)
        Major(i + 1) = Temp
        Temp = Hometown(i)
        Hometown(i) = Hometown(i + 1)
        Hometown(i + 1) = Temp
        Temp = Birthdate(i)
        Birthdate(i) = Birthdate(i + 1)
        Birthdate(i + 1) = Temp
        End If
    Next i
Next Pass

For i = 1 To N
pbxResults.Print RoomNumber(i); Tab(25); strName(i); Tab(55); Class(i); Tab(65); Major(i); Tab(75); Hometown(i); Tab(100); Birthdate(i)
Next i
Close #1
End Sub

