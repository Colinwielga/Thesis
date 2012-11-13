VERSION 5.00
Begin VB.Form AircraftSpecs 
   BackColor       =   &H00FF8080&
   Caption         =   "Aircraft Specifications"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form2"
   ScaleHeight     =   7860
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmainmenu 
      Caption         =   "Return To Main Menu"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      TabIndex        =   6
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CommandButton cmdlookup 
      Caption         =   "Look Up Aircraft"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox pbxresults1 
      Height          =   5895
      Left            =   2040
      ScaleHeight     =   5835
      ScaleWidth      =   8715
      TabIndex        =   4
      Top             =   240
      Width           =   8775
   End
   Begin VB.CommandButton cmdsortrange 
      Caption         =   "Sort By Aircraft Range"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdsortspeed 
      Caption         =   "Sort By Cruise Speed"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdsortpass 
      Caption         =   "Sort By Passenger Load"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdreadprint 
      Caption         =   "Read And Print Technical Specifications"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblauthor 
      BackColor       =   &H00FF8080&
      Caption         =   "VB Design by Kerry R. O'Neill 10/24/2003"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label lblweight 
      BackColor       =   &H00FF8080&
      Caption         =   "Weight is maximum allowable for takeoff"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Label lblnm 
      BackColor       =   &H00FF8080&
      Caption         =   "Range Indicated in Nautical Miles"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label lblspeed 
      BackColor       =   &H00FF8080&
      Caption         =   "Speed Indicated in MPH"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   6360
      Width           =   2895
   End
End
Attribute VB_Name = "AircraftSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this form is to provide the user with
'information about Boeing's current product line. This
'includes pertinent information about each aircraft's
'performance and characteristics.


Option Explicit
Dim N(1 To 8) As String 'Aircraft Name
Dim P(1 To 8) As Single '# passengers, 2 class
Dim C(1 To 8) As Single 'pounds cargo capacity
Dim E(1 To 8) As Single '# of engines
Dim T(1 To 8) As Single 'Engine Thrust
Dim W(1 To 8) As Single 'Aircraft Weight (max.)
Dim R(1 To 8) As Single 'Maximum Range
Dim S(1 To 8) As Single 'Typical Cruising Speed
Dim i As Integer
Public strpath As String



Private Sub cmdmainmenu_Click() 'returns user to the main page
    AircraftSpecs.Hide
    MainMenu.Show
End Sub

Private Sub cmdlookup_Click() 'looks up an aircraft after the users inputs the aircraft name
    pbxresults1.Cls
    Dim Y As String
    Dim done As Boolean
    Y = InputBox("Enter Name", "Aircraft Search")
    done = False
    i = 0
    Do Until done Or i = 8  'searches until match is found
       i = i + 1 'moves onto next name
       If Y = N(i) Then done = True
    Loop
    If done Then 'prints what is found
        pbxresults1.Print "Name", Tab(15); "Passengers", Tab(30); "Cargo(lbs)", Tab(45); "Engines", Tab(59); "Thrust(lbs)", Tab(75); "Weight", Tab(90); "Range", Tab(107); "Speed"
        pbxresults1.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        pbxresults1.Print 'provides a space between heading and data
        pbxresults1.Print N(i), Tab(15); P(i), Tab(30); C(i), Tab(45); E(i), Tab(59); T(i), Tab(75); W(i), Tab(90); R(i), Tab(107); S(i)
    Else
        MsgBox ("Aircraft Not Found") 'alerts user that entry is not valid
    End If
    
End Sub

Private Sub cmdreadprint_Click() 'reads text file and inputs it into vb program
    
    Open strpath & "boeingspecs.txt" For Input As #1 'opens text file from folder
    pbxresults1.Cls
    pbxresults1.Print "Name", Tab(15); "Passengers", Tab(30); "Cargo(lbs)", Tab(45); "Engines", Tab(59); "Thrust(lbs)", Tab(75); "Weight", Tab(90); "Range", Tab(107); "Speed"
    pbxresults1.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    pbxresults1.Print 'provides a space between heading and data
    For i = 1 To 8
        Input #1, N(i), P(i), C(i), E(i), T(i), W(i), R(i), S(i) 'loads file into 8 parallel arrays
        pbxresults1.Print N(i), Tab(15); P(i), Tab(30); C(i), Tab(45); E(i), Tab(59); T(i), Tab(75); W(i), Tab(90); R(i), Tab(107); S(i) 'prints information onto table in the picture box
    Next i
    Close #1
    
End Sub

Private Sub cmdsortpass_Click() 'displays each aircraft by maximum passenger complement
    Dim pass As Integer
    Dim X As Integer
    Dim temp1 As String 'temp files for use in sorting sequences
    Dim temp2 As Single
    Dim temp3 As Single
    Dim temp4 As Single
    Dim temp5 As Single
    Dim temp6 As Single
    Dim temp7 As Single
    Dim temp8 As Single
    X = 8
    pbxresults1.Cls
    pbxresults1.Print "Name", Tab(15); "Passengers", Tab(30); "Cargo(lbs)", Tab(45); "Engines", Tab(59); "Thrust(lbs)", Tab(75); "Weight", Tab(90); "Range", Tab(107); "Speed"
    pbxresults1.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    pbxresults1.Print 'provides a space between heading and data
    For pass = 1 To (X - 1) 'bubble sort to sort information by passenger payload in descending order
        For i = 1 To (X - pass)
            If P(i) < P(i + 1) Then
                temp1 = N(i)
                N(i) = N(i + 1)
                N(i + 1) = temp1
                temp2 = P(i)
                P(i) = P(i + 1)
                P(i + 1) = temp2
                temp3 = C(i)
                C(i) = C(i + 1)
                C(i + 1) = temp3
                temp4 = E(i)
                E(i) = E(i + 1)
                E(i + 1) = temp4
                temp5 = T(i)
                T(i) = T(i + 1)
                T(i + 1) = temp5
                temp6 = W(i)
                W(i) = W(i + 1)
                W(i + 1) = temp6
                temp7 = R(i)
                R(i) = R(i + 1)
                R(i + 1) = temp7
                temp8 = S(i)
                S(i) = S(i + 1)
                S(i + 1) = temp8
            End If
        Next i
    Next pass
    
    For i = 1 To 8
        pbxresults1.Print N(i), Tab(15); P(i), Tab(30); C(i), Tab(45); E(i), Tab(59); T(i), Tab(75); W(i), Tab(90); R(i), Tab(107); S(i) 'prints sorted information
    Next i
    
End Sub

Private Sub cmdsortrange_Click() 'sorts aircraft by maximum range
    Dim pass As Integer
    Dim X As Integer
    Dim temp1 As String 'temp files for use in sorting sequences
    Dim temp2 As Single
    Dim temp3 As Single
    Dim temp4 As Single
    Dim temp5 As Single
    Dim temp6 As Single
    Dim temp7 As Single
    Dim temp8 As Single
    X = 8
    pbxresults1.Cls
    pbxresults1.Print "Name", Tab(15); "Passengers", Tab(30); "Cargo(lbs)", Tab(45); "Engines", Tab(59); "Thrust(lbs)", Tab(75); "Weight", Tab(90); "Range", Tab(107); "Speed"
    pbxresults1.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    pbxresults1.Print 'provides a space between heading and data
    For pass = 1 To (X - 1) 'bubble sort to sort information by range in descending order
        For i = 1 To (X - pass)
            If R(i) < R(i + 1) Then
                temp1 = N(i)
                N(i) = N(i + 1)
                N(i + 1) = temp1
                temp2 = P(i)
                P(i) = P(i + 1)
                P(i + 1) = temp2
                temp3 = C(i)
                C(i) = C(i + 1)
                C(i + 1) = temp3
                temp4 = E(i)
                E(i) = E(i + 1)
                E(i + 1) = temp4
                temp5 = T(i)
                T(i) = T(i + 1)
                T(i + 1) = temp5
                temp6 = W(i)
                W(i) = W(i + 1)
                W(i + 1) = temp6
                temp7 = R(i)
                R(i) = R(i + 1)
                R(i + 1) = temp7
                temp8 = S(i)
                S(i) = S(i + 1)
                S(i + 1) = temp8
            End If
        Next i
    Next pass
    
    For i = 1 To 8
        pbxresults1.Print N(i), Tab(15); P(i), Tab(30); C(i), Tab(45); E(i), Tab(59); T(i), Tab(75); W(i), Tab(90); R(i), Tab(107); S(i) 'prints sorted information
    Next i
End Sub
Private Sub cmdsortspeed_Click() 'Sorts by Cruising speed of Aircraft(incidentally the same as passenger load)

    Dim pass As Integer
    Dim X As Integer
    Dim temp1 As String 'temp files for use in sorting sequences
    Dim temp2 As Single
    Dim temp3 As Single
    Dim temp4 As Single
    Dim temp5 As Single
    Dim temp6 As Single
    Dim temp7 As Single
    Dim temp8 As Single
    X = 8
    pbxresults1.Cls
    pbxresults1.Print "Name", Tab(15); "Passengers", Tab(30); "Cargo(lbs)", Tab(45); "Engines", Tab(59); "Thrust(lbs)", Tab(75); "Weight", Tab(90); "Range", Tab(107); "Speed"
    pbxresults1.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    pbxresults1.Print 'provides a space between heading and data
    For pass = 1 To (X - 1) 'bubble sort to sort the information by cruising speed in descending order
        For i = 1 To (X - pass)
            If S(i) < S(i + 1) Then
                temp1 = N(i)
                N(i) = N(i + 1)
                N(i + 1) = temp1
                temp2 = P(i)
                P(i) = P(i + 1)
                P(i + 1) = temp2
                temp3 = C(i)
                C(i) = C(i + 1)
                C(i + 1) = temp3
                temp4 = E(i)
                E(i) = E(i + 1)
                E(i + 1) = temp4
                temp5 = T(i)
                T(i) = T(i + 1)
                T(i + 1) = temp5
                temp6 = W(i)
                W(i) = W(i + 1)
                W(i + 1) = temp6
                temp7 = R(i)
                R(i) = R(i + 1)
                R(i + 1) = temp7
                temp8 = S(i)
                S(i) = S(i + 1)
                S(i + 1) = temp8
            End If
        Next i
    Next pass
    
    For i = 1 To 8
        pbxresults1.Print N(i), Tab(15); P(i), Tab(30); C(i), Tab(45); E(i), Tab(59); T(i), Tab(75); W(i), Tab(90); R(i), Tab(107); S(i) 'prints sorted information
    Next i
End Sub


Private Sub Form_Load() 'creates a strpath so the file can be opened after being moved to different folders
    strpath = "N:\CS130\handin\KRONEILL\"
End Sub
