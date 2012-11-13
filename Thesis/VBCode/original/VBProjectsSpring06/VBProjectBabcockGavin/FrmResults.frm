VERSION 5.00
Begin VB.Form frmResults 
   BackColor       =   &H00400040&
   Caption         =   "Results Page"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11520
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H00FFC0C0&
      Caption         =   "About the Authors"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8760
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   8520
      TabIndex        =   8
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton cmdAllConference 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show All-Conference Runners"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CommandButton cmdNavigate3 
      BackColor       =   &H000080FF&
      Caption         =   "Go To Statistics Page"
      BeginProperty Font 
         Name            =   "JazzTextExtended"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8880
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   6975
      Left            =   2640
      ScaleHeight     =   6915
      ScaleWidth      =   5475
      TabIndex        =   5
      Top             =   1560
      Width           =   5535
   End
   Begin VB.CommandButton cmdNavigate2 
      BackColor       =   &H000080FF&
      Caption         =   "Go To Splits Page"
      BeginProperty Font 
         Name            =   "JazzTextExtended"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8880
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortYear 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort Results by Year in School"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortSchool 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort Results by School"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort Results Alphabetically"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoadResults 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Load Race Results by Time"
      BeginProperty Font 
         Name            =   "Eras Light ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Image ImgSweetness 
      Height          =   6750
      Left            =   8520
      Picture         =   "FrmResults.frx":0000
      Top             =   1680
      Width           =   1980
   End
   Begin VB.Label lblNames 
      BackColor       =   &H00400040&
      Caption         =   $"FrmResults.frx":2B85A
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   10560
      Width           =   10335
   End
   Begin VB.Image imgTitle 
      Height          =   1455
      Left            =   1920
      Picture         =   "FrmResults.frx":2B8E8
      Top             =   0
      Width           =   6555
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pos As Integer
'The purpose of this program is to create a beneficial application
'to calculate and organize race results from a 5k race


Private Sub cmdAbout_Click()
    frmAuthors.Show
    frmResults.Hide
End Sub

Private Sub cmdAllConference_Click()
    picResults.Cls
    Pos = 0
    picResults.Print Tab(20); " All Conference Runners "
    picResults.Print
    'Print the header info
    picResults.Print "Rank ", "Name", "Year", "School", "  Time"
    For Pos = 1 To ArraySize
        Select Case Minutes(Pos)
            Case Is <= 15
                picResults.Print Tab(2); Place(Pos); Tab(10); Names(Pos); Tab(30); Year(Pos); Tab(43); School(Pos); Tab(57); Minutes(Pos); ":"; Seconds(Pos)
        End Select
    Next Pos
End Sub

Private Sub cmdLoadResults_Click()
        
    Call PlaySound("mywav.wav")
    
    'clear the picture screen
    picResults.Cls

    'set counter (pos) at zero
    Pos = 0
    'Open the file
    Open App.Path & "\RunnerResults.txt" For Input As #1
    'Print the header info
    picResults.Print "Place", "Name", "Year", "School", "  Time"
    'begin loop to load into the array and print each line in the text file


    Do While Not EOF(1)
        Pos = Pos + 1
        'write the array with the given information
        Input #1, Place(Pos), Names(Pos), Year(Pos), School(Pos), Minutes(Pos), Seconds(Pos)
        'print and format the results one at a time
        picResults.Print Tab(2); Place(Pos); Tab(10); Names(Pos); Tab(30); Year(Pos); Tab(43); School(Pos); Tab(57); Minutes(Pos); ":"; Seconds(Pos)
    Loop
    
    'Make ArraySize = Pos so that we know how large the array is for future commands.
    ArraySize = Pos
    
    'close the file
    Close #1

End Sub

Private Sub cmdNavigate2_Click()
    frmSplits.Show
    frmResults.Hide
    
End Sub

Private Sub cmdNavigate3_Click()
    frmStats.Show
    frmResults.Hide
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSortName_Click()
    Dim PlaceTemp As Integer
    Dim NameTemp As String
    Dim YearTemp As Integer
    Dim SchoolTemp As String
    Dim MinutesTemp As Single
    Dim SecondsTemp As Single
    Dim Pass As Integer
    
    
    'this clears the screen
    picResults.Cls
    'This prints the header information
    picResults.Print Tab(2); "Place"; Tab(10); "Name"; Tab(30); "Year"; Tab(43); "School"; Tab(57); "Time"
    
    For Pass = 1 To ArraySize - 1
        For Pos = 1 To ArraySize - Pass
            If Names(Pos) > Names(Pos + 1) Then
                'organize and replace the temp variable if it meets the given requirements
                MinutesTemp = Minutes(Pos)
                Minutes(Pos) = Minutes(Pos + 1)
                Minutes(Pos + 1) = MinutesTemp
                
                SecondsTemp = Seconds(Pos)
                Seconds(Pos) = Seconds(Pos + 1)
                Seconds(Pos + 1) = SecondsTemp
                
                PlaceTemp = Place(Pos)
                Place(Pos) = Place(Pos + 1)
                Place(Pos + 1) = PlaceTemp
                
                NameTemp = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = NameTemp
                
                YearTemp = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = YearTemp
                
                SchoolTemp = School(Pos)
                School(Pos) = School(Pos + 1)
                School(Pos + 1) = SchoolTemp
            End If
        Next Pos
    Next Pass
        'print the information using a for-next loop
        For Pos = 1 To ArraySize
            picResults.Print Tab(2); Place(Pos); Tab(10); Names(Pos); Tab(30); Year(Pos); Tab(43); School(Pos); Tab(57); Minutes(Pos); ":"; Seconds(Pos)
       Next Pos
End Sub

Private Sub cmdSortSchool_Click()
    Dim PlaceTemp As Integer
    Dim NameTemp As String
    Dim YearTemp As Integer
    Dim SchoolTemp As String
    Dim MinutesTemp As Single
    Dim SecondsTemp As Single
    Dim Pass As Integer
    
    

   
    picResults.Cls
    picResults.Print Tab(2); "Place"; Tab(10); "Name"; Tab(30); "Year"; Tab(43); "School"; Tab(57); "Time"
    
    For Pass = 1 To ArraySize - 1
        For Pos = 1 To ArraySize - Pass
            If School(Pos) > School(Pos + 1) Then
                MinutesTemp = Minutes(Pos)
                Minutes(Pos) = Minutes(Pos + 1)
                Minutes(Pos + 1) = MinutesTemp
                
                SecondsTemp = Seconds(Pos)
                Seconds(Pos) = Seconds(Pos + 1)
                Seconds(Pos + 1) = SecondsTemp

                PlaceTemp = Place(Pos)
                Place(Pos) = Place(Pos + 1)
                Place(Pos + 1) = PlaceTemp
                
                NameTemp = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = NameTemp
                
                YearTemp = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = YearTemp
                
                SchoolTemp = School(Pos)
                School(Pos) = School(Pos + 1)
                School(Pos + 1) = SchoolTemp
            End If
        Next Pos
    Next Pass
        
        For Pos = 1 To ArraySize
            picResults.Print Tab(2); Place(Pos); Tab(10); Names(Pos); Tab(30); Year(Pos); Tab(43); School(Pos); Tab(57); Minutes(Pos); ":"; Seconds(Pos)
        Next Pos
    End Sub
    



Private Sub cmdSortYear_Click()
    Dim PlaceTemp As Integer
    Dim NameTemp As String
    Dim YearTemp As Integer
    Dim SchoolTemp As String
    Dim MinutesTemp As Single
    Dim SecondsTemp As Single
    Dim Pass As Integer
    
    

    picResults.Cls
    picResults.Print Tab(2); "Place"; Tab(10); "Name"; Tab(30); "Year"; Tab(43); "School"; Tab(57); "Time"
    
    For Pass = 1 To ArraySize - 1
        For Pos = 1 To ArraySize - Pass
            If Year(Pos) > Year(Pos + 1) Then
                MinutesTemp = Minutes(Pos)
                Minutes(Pos) = Minutes(Pos + 1)
                Minutes(Pos + 1) = MinutesTemp
                
                
                SecondsTemp = Seconds(Pos)
                Seconds(Pos) = Seconds(Pos + 1)
                Seconds(Pos + 1) = SecondsTemp
                
                PlaceTemp = Place(Pos)
                Place(Pos) = Place(Pos + 1)
                Place(Pos + 1) = PlaceTemp
                
                NameTemp = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = NameTemp
                
                YearTemp = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = YearTemp
                
                SchoolTemp = School(Pos)
                School(Pos) = School(Pos + 1)
                School(Pos + 1) = SchoolTemp
            End If
        Next Pos
    Next Pass
    For Pos = 1 To ArraySize
        picResults.Print Tab(2); Place(Pos); Tab(10); Names(Pos); Tab(30); Year(Pos); Tab(43); School(Pos); Tab(57); Minutes(Pos); ":"; Seconds(Pos)
    Next Pos
End Sub


