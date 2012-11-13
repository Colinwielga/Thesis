VERSION 5.00
Begin VB.Form frmSkiResults 
   BackColor       =   &H00FF0000&
   Caption         =   "Nordic Ski Results Program"
   ClientHeight    =   10500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15840
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   15840
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptManmadeSnow 
      BackColor       =   &H000040C0&
      Caption         =   "Manmade Snow"
      Height          =   615
      Left            =   4080
      TabIndex        =   11
      Top             =   7080
      Width           =   1215
   End
   Begin VB.OptionButton OptOldSnow 
      BackColor       =   &H000080FF&
      Caption         =   "Old crusty snow"
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.OptionButton OptNewSnow 
      BackColor       =   &H0080C0FF&
      Caption         =   "New Snow"
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   7080
      Width           =   1215
   End
   Begin VB.PictureBox picImage2 
      Height          =   2775
      Left            =   11640
      Picture         =   "frmSkiResults.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   4035
      TabIndex        =   7
      Top             =   1800
      Width           =   4095
   End
   Begin VB.PictureBox picImage1 
      Height          =   3615
      Left            =   240
      Picture         =   "frmSkiResults.frx":36AC
      ScaleHeight     =   3555
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   1440
      Width           =   5895
   End
   Begin VB.CommandButton cmdWaxTempConv 
      BackColor       =   &H0000FFFF&
      Caption         =   "Wax Tempurature Conversion"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8400
      Width           =   3255
   End
   Begin VB.CommandButton cmdShowClassic 
      BackColor       =   &H000040C0&
      Caption         =   "Proceed to Classic Results =======>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSTimes 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter Skate Times"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   3615
   End
   Begin VB.CommandButton cmdCTimes 
      BackColor       =   &H000080FF&
      Caption         =   "Enter Classic Times"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6960
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label lblWax 
      BackColor       =   &H00FF0000&
      Caption         =   "For accurate wax advice please check the corresponding box and then click on the conversion button below."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   840
      TabIndex        =   8
      Top             =   5520
      Width           =   4695
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FF0000&
      Caption         =   "Nordic Ski Results Program"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   11175
   End
End
Attribute VB_Name = "frmSkiResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program takes times from nordic ski races enter by the user and sorts them to determine winner and winning school.
'There are two different races and then the times are added together to have an overall winner.

Private Sub cmdCTimes_Click()
    'Matches times inputed by the user to data already created in Access
    Dim CTR As Integer
    'Dim variables
    Dim DB As Database
    Dim RS As Recordset2
    Dim S As Integer
    
    'Open database and recordset
    Set DB = OpenDatabase(App.Path & "\SkiRaceResults1.accdb")
    Set RS = DB.OpenRecordset("SkiRace1")
    RS.Index = "PrimaryKey"
   
   'Puts data into two arrays
    Do Until RS.EOF
        pos = pos + 1
        Bib(pos) = RS![Bib]
        'picResults.Print Bib(pos)
        SkierFName(pos) = RS![First Name]
        'picResults.Print SkierFName(pos)
        SkierLName(pos) = RS![Last Name]
        'picResults.Print SkierLName(pos)
        School(pos) = RS![School]
        'picResults.Print School(pos)
        CTimes(pos) = InputBox("Enter a Classic time (hr:min:sec) for Bib " & RS![Bib] & ".")
        'picResults.Print CTimes(pos)
        'STimes(pos) = InputBox("Enter a Skate time (hr:min:sec) for Bib " & RS![Bib] & ".")
        RS.MoveNext
    Loop
    
    
    'picResults.Print "*********************"
    'picResults.Print "pos is "; pos
    'For S = 1 To pos
        'picResults.Print School(S)
        'picResults.Print SkierFName(S), SkierLName(S), Bib(S), School(S), Minute(CTimes(S)), Second(CTimes(S))
    'Next S
    'Close file
    RS.Close
    DB.Close
        
End Sub

Private Sub cmdNewRacer_Click()
    Dim P As Integer
    P = pos + 1
    
    For P = 1 To 1
        SkierFName(P) = InputBox("Enter the skier's first name: ")
        SkierLName(P) = InputBox("Enter the skier's last name:")
        Bib(P) = InputBox("Enter the skier's bib number:")
        School(P) = InputBox("Enter the skier's school:")
        CTimes(P) = InputBox("Enter the skier's clsssic time:")
        STimes(P) = InputBox("Enter the skier's skate time:")
        
    Next P
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdShowClassic_Click()
    'Changes forms
    frmClassic.Show
    frmSkiResults.Hide
End Sub

Private Sub cmdSTimes_Click()
    Dim DB As Database
    Dim RS As Recordset2
    Set DB = OpenDatabase(App.Path & "\SkiRaceResults1.accdb")
    Set RS = DB.OpenRecordset("SkiRace1")
    RS.Index = "PrimaryKey"
    
    Do Until RS.EOF
        pos1 = pos1 + 1
        'Bib(pos1) = RS![Bib]
        SkierFName(pos1) = RS![First Name]
        SkierLName(pos1) = RS![Last Name]
        School(pos1) = RS![School]
        
        STimes(pos1) = InputBox("Enter a skate time (hr:min:sec) for Bib " & RS![Bib] & ".")
        RS.MoveNext
    Loop
    RS.Close
    DB.Close
        
   'For CTR = 1 To pos1
        'picResults.Print Bib(CTR), STimes(CTR)
   'Next CTR
End Sub

Private Sub cmdWaxTempConv_Click()
    'Determines what wax the racers should have used during race based on option button and temperature inputed by the user.
    Dim snow As String
    Dim NewSnow As String
    Dim OldSnow As String
    Dim Manmade As String
    Dim temperature As Single
   
    If OptNewSnow.Value = False _
        And OptOldSnow.Value = False Then
        MsgBox "Click on a snow condition"
    ElseIf OptNewSnow.Value = True Then
        snow = NewSnow
    ElseIf OptOldSnow.Value = True Then
        snow = OldSnow
    Else
        snow = Manmade
    End If

    temperature = InputBox("Enter the temperature (in farenheit) at race time to determine which type of Swix wax you should use.", "Temperature")
    
    Select Case temperature
        Case Is >= 40
            If snow = NewSnow Then
                MsgBox "It's a hot one! I can't believe there is still snow left! Use LF10 or HF10 (Yellow)"
            ElseIf snow = OldSnow Then
                MsgBox "It's a hot one! I can't believe there is still snow left! Use LF10 or HF10 (Yellow)"
            Else
                MsgBox "It's a hot one! I can't believe there is still snow left! Use LF10 or HF10 (Yellow)"
            End If
           
        Case 32 To 39
            If snow = NewSnow Then
                MsgBox "It's above freezing! Spring is coming. Use LF8 or HF8 (Red)"
                
            ElseIf snow = OldSnow Then
                 MsgBox "It's above freezing! Spring is coming. Use LF8 or HF8 (Red)"
            Else
                 MsgBox "It's above freezing! Spring is coming. Use LF8 or HF7 (Red)"
            End If
        Case 20 To 31
            If snow = NewSnow Then
                 MsgBox "It's a great day to be out skiing today! Use LF8 or HF8 (Red)"
            ElseIf snow = OldSnow Then
                 MsgBox "It's a great day to be out skiing today! Use LF7 or LF8 (Red or Purple)"
            Else
                 MsgBox "It's a great day to be out skiing today! Use LF7 (Purple)"
            End If
        Case 10 To 19
            If snow = NewSnow Then
                  MsgBox "It's cold but once you get skiing it will be great. Use LF6 or HF6 (Blue)"
            ElseIf snow = OldSnow Then
                 MsgBox "It's cold but once you get skiing it will be great. Use LF6 or HF6 (Blue)"
            Else
                 MsgBox "It's cold but once you get skiing it will be great. Use LF6 or HF6 (Blue)"
            End If
        Case 0 To 9
            If snow = NewSnow Then
                  MsgBox "It's chilly today! Use LF4 or HF4 (Green)"
            ElseIf snow = OldSnow Then
                MsgBox "It's chilly today! Use LF4 or HF4 (Green)"
            Else
                 MsgBox "It's chilly today! Use LF4 or HF4 (Green)"
            End If
        Case -10 To -1
            If snow = NewSnow Then
                 MsgBox "It's cold today, dress warmly. Use CH4 (Green)"
            ElseIf snow = OldSnow Then
                MsgBox "It's cold today, dress warmly. Use CH4 (Green)"
            Else
                 MsgBox "It's cold today, dress warmly. Use CH4 (Green)"
            End If
        Case Else
            MsgBox "Don't ski today! You should go back inside!"
        End Select
End Sub

Private Sub Form1_Click()
    OptNewSnow.Value = False
    OptOldSnow.Value = False
    OptManmadeSnow.Value = False
    
End Sub
