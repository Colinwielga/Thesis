VERSION 5.00
Begin VB.Form frmCountryInfo 
   BackColor       =   &H80000009&
   Caption         =   "Discover Western Europe - Country Information"
   ClientHeight    =   7650
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   17145
   LinkTopic       =   "Form3"
   ScaleHeight     =   7650
   ScaleWidth      =   17145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUKtravel 
      Caption         =   "United Kingdom Travel Guide"
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdSUItravel 
      Caption         =   "Switzerland Travel Guide"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdPORtravel 
      Caption         =   "Portugal Travel Guide"
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdNEDtravel 
      Caption         =   "Netherlands Travel Guide"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdITAtravel 
      Caption         =   "Italy Travel Guide"
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdIREtravel 
      Caption         =   "Ireland Travel Guide"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdGERtravel 
      Caption         =   "Germany Travel Guide"
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdFRAtravel 
      Caption         =   "France Travel Guide"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdESPtravel 
      Caption         =   "Spain Travel Guide"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdBELtravel 
      Caption         =   "Belgium Travel Guide"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdUKMAP 
      Caption         =   "United Kingdom Detailed Map"
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdSUIMAP 
      Caption         =   "Switzerland Detailed Map"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdPORMAP 
      Caption         =   "Portugal Detailed Map"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdESPMAP 
      Caption         =   "Spain Detailed Map"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdNEDMAP 
      Caption         =   "Netherlands Detailed Map"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdITAMAP 
      Caption         =   "Italy Detailed Map"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdIREMAP 
      Caption         =   "Ireland Detailed Map"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdGERMAP 
      Caption         =   "Germany Detailed Map"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdFRAMAP 
      Caption         =   "France Detailed Map"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdBELMAP 
      Caption         =   "Belgium Detailed Map"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdvisible 
      Caption         =   "Return to Country List"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdOthers 
      Caption         =   "What about those other Western European Countries?!"
      Height          =   375
      Left            =   8880
      TabIndex        =   19
      Top             =   960
      Width           =   8055
   End
   Begin VB.CommandButton cmdABCDesending 
      Caption         =   "Alphabetize by City Desending"
      Enabled         =   0   'False
      Height          =   495
      Left            =   15000
      TabIndex        =   18
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdABCAsending 
      Caption         =   "Alphabetize by City Asending"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12960
      TabIndex        =   17
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdCitiesDesending 
      Caption         =   "Sort City Population Desending"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10920
      TabIndex        =   16
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdCitiesAsending 
      Caption         =   "Sort City Population Asending"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8880
      TabIndex        =   15
      Top             =   1440
      Width           =   1935
   End
   Begin VB.PictureBox picFlag 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   2655
      TabIndex        =   14
      Top             =   1080
      Width           =   2655
   End
   Begin VB.PictureBox picMap 
      Height          =   6375
      Left            =   2400
      ScaleHeight     =   6315
      ScaleWidth      =   6315
      TabIndex        =   12
      Top             =   960
      Width           =   6375
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdUKInfo 
      Caption         =   "United Kingdom"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdSUIInfo 
      Caption         =   "Switzerland"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdESPInfo 
      Caption         =   "Spain"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdPORInfo 
      Caption         =   "Portugal"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton cmdNEDInfo 
      BackColor       =   &H0000C000&
      Caption         =   "Netherlands"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdITAInfo 
      Caption         =   "Italy"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdIREInfo 
      Caption         =   "Ireland"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.PictureBox picCities 
      Height          =   5415
      Left            =   8880
      ScaleHeight     =   5355
      ScaleWidth      =   7995
      TabIndex        =   3
      Top             =   2040
      Width           =   8055
   End
   Begin VB.CommandButton cmdGERInfo 
      Caption         =   "Germany"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdFRAInfo 
      Caption         =   "France"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdBELInfo 
      Caption         =   "Belgium"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H80000009&
      Caption         =   "Click on the a button to see Country Info and Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmCountryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmCountryInfo (Discover Western Europe)
'Author: Nate Burbeck (map buttons by Brad Hagemeier)
'Date Written: 23 March 2008
'Objective: This form is the information section of the project. The user can click a button for each of the ten countries featured in this project which will show that country on
'the Europe map along with its flag and data on that country's most populous cities.  The user can sort the list of cities alphabetically or by population.  The user will also have the
'option of viewing a detailed map of that country and well as viewing that countries wikitravel page.
Option Explicit
Dim CTR As Integer, Metro(1 To 100) As String, Population(1 To 100) As Long, Region(1 To 100) As String
Dim UKActive As Boolean, FRAActive As Boolean, BELActive As Boolean, ESPActive As Boolean, GERActive As Boolean
Dim IREActive As Boolean, ITAActive As Boolean, NEDActive As Boolean, PORActive As Boolean, SUIActive As Boolean
'sets all variables and arrays
'enables linking to a web address

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'These may seem incredibly long but really it's rather redundant, each section is repeated for each country.
'When say cmdUKInfo is pressed the program reads the corrisponding array (UKcities.txt), enables 4 sorting buttons (cmdCitiesAsending, cmdCitiesDesending, cmdABCAsending, cmdABCDesending),
'and makes visible 3 other buttons (cmdUKMap, cmdvisible, and cmdUKtravel)
'a picture is also displayed with the country's location in europe and a flag
'This is repeated for each of the 10 countries

Private Sub cmdABCAsending_Click()              'this button is only made active when a country info button is pressed (i.e. cmdUKInfo)
If UKActive = True Then                         'UKActive was set to true when cmdUKInfo was pressed, thus it will only sort by the array that
        Dim Pass As Integer                     'has been read when cmdUKInfo was pressed (that would be UKcities.txt)
        Dim Pos As Integer
        Dim TempPop As Long
        Dim TempRegion As String
        Dim TempMetro As String
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then         'sorts the array that has already been read when cmdUKInfo was pressed
                    TempPop = Population(Pos)               'the string array 'Metro' is used as the anchor for this bubble sort
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        Dim i As Integer
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If FRAActive = True Then
        For Pass = 1 To CTR - 1                                         'now the same for France (when cmdFRAInfo is pressed FRAActive is set to True)
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If BELActive = True Then
        For Pass = 1 To CTR - 1                                 'and Belgium
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If ESPActive = True Then
        For Pass = 1 To CTR - 1                             'and all the other countries
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If GERActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If IREActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If ITAActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If NEDActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If PORActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If SUIActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) > Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
End Sub

Private Sub cmdABCDesending_Click()
If UKActive = True Then
        Dim Pass As Integer
        Dim Pos As Integer
        Dim TempPop As Long
        Dim TempRegion As String
        Dim TempMetro As String
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        Dim i As Integer
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If FRAActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If BELActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If ESPActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If GERActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If IREActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If ITAActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If NEDActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If PORActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If SUIActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Metro(Pos) < Metro(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
End Sub

Private Sub cmdBELInfo_Click()          'code is similar for each country, enabling different options
picCities.Cls                           'picture box is cleared so that BELcities.txt can be displayed
cmdCitiesAsending.Enabled = True        'enables sorting buttons
cmdCitiesDesending.Enabled = True       'ditto
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
cmdBELtravel.Visible = True             'makes cmdBELtravel visible which will be used to link to a web page
BELActive = True                        'sets BELActive to True, this will be used to bubble sort the cities array file BELcities.txt (see above code)
picMap.Picture = LoadPicture(App.Path & "\Images\mapBEL.JPG")       'loads the map of belgium
picFlag.Picture = LoadPicture(App.Path & "\Images\flagBEL.GIF")     'loads the belgian flag which is superimposed ontop of the map
CTR = 0                                                     'CTR set to 0
    Open App.Path & "\BELcities.txt" For Input As #1        'reads BELcities as an array
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)  'reads file into Metro, Population and Region
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)    'prints BELcities.txt as arrays
    Next i                                                                                       'loops back
cmdGERInfo.Visible = False      'sets germany button to invisible
cmdIREInfo.Visible = False      'same for ireland
cmdESPInfo.Visible = False      'and spain
cmdFRAInfo.Visible = False      'and so on
cmdBELInfo.Visible = False      'and so forth
cmdITAInfo.Visible = False      'yada yada yada
cmdNEDInfo.Visible = False      'you get the picture
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True       'sets cmdvisible to visible (this is the button that says 'Return to country list')
cmdBELMAP.Visible = True        'BELMAP is set to visible


End Sub

Private Sub cmdBELMAP_Click()       'when cmdBELInfo is clicked cmdBELMAP is made visible
    frmBELMAP.Visible = True        'shows frmBELMAP, a detailed map of Belgium
    frmCountryInfo.Visible = False  'previous form, frmCountryInfo, is changed to invisible
End Sub

Private Sub cmdCitiesAsending_Click()   'sorts arrays by City population Asending, notice that each array case (country, UKcities for example) is enabled
If UKActive = True Then                 'only when its corrisponding Info button is pressed and it is made active (UKActive = True) because this is the currently read array
        Dim Pass As Integer             'if UKActive is true then cmdCitiesAsending will only sort arrays in UKcities.txt, the same goes for
        Dim Pos As Integer              'other countries further down the code in this button
        Dim TempPop As Long
        Dim TempRegion As String
        Dim TempMetro As String
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then       'bubble sorting is anchored by population array
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        Dim i As Integer
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If FRAActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If BELActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If ESPActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If GERActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If IREActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If ITAActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If NEDActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If PORActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If SUIActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
End Sub

Private Sub cmdCitiesDesending_Click()                   'same code as cmdCitiesAsending but now it decends by the array 'Population'
If UKActive = True Then
        Dim Pass As Integer
        Dim Pos As Integer
        Dim TempPop As Long
        Dim TempRegion As String
        Dim TempMetro As String
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) < Population(Pos + 1) Then       'sorting is anchored by the population array
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro                          'isn't it beautiful?
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        Dim i As Integer
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If FRAActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) > Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If BELActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) < Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If ESPActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) < Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If GERActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) < Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If IREActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) < Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If ITAActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) < Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If NEDActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) < Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If PORActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) < Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)                           'almost there
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
If SUIActive = True Then
        For Pass = 1 To CTR - 1
            For Pos = 1 To CTR - Pass
                If Population(Pos) < Population(Pos + 1) Then
                    TempPop = Population(Pos)
                    Population(Pos) = Population(Pos + 1)
                    Population(Pos + 1) = TempPop
                    TempRegion = Region(Pos)
                    Region(Pos) = Region(Pos + 1)
                    Region(Pos + 1) = TempRegion
                    TempMetro = Metro(Pos)
                    Metro(Pos) = Metro(Pos + 1)
                    Metro(Pos + 1) = TempMetro
                End If
            Next Pos
        Next Pass
        picCities.Cls
        picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
        picCities.Print "------------------------------------------------------------------------------------------------------------"
        For i = 1 To CTR
            picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
        Next i
    End If
End Sub

Private Sub cmdESPMAP_Click()       'similar to cmdBELMAP above, only the name is switched, in the form each
frmCountryInfo.Hide                 'of these buttons are stacked ontop of eachother but only made visible when
frmESPMAP.Show                      'the corrisponding country info button is pressed by the user
End Sub

Private Sub cmdFRAMAP_Click()
frmCountryInfo.Hide
frmFRAMAP.Show
End Sub

Private Sub cmdGERMAP_Click()
frmCountryInfo.Hide
frmGERMAP.Show
End Sub

Private Sub cmdIREMAP_Click()
frmCountryInfo.Hide
frmIREMAP.Show
End Sub

Private Sub cmdITAMAP_Click()
frmCountryInfo.Hide
frmITAMAP.Show
End Sub

Private Sub cmdMainMenu_Click()

frmMainMenu.Show                        'returns to the main menu, this form is made invisible
frmCountryInfo.Hide

End Sub

Private Sub cmdmap_Click()
    frmBELMAP.Show
    frmCountryInfo.Hide
End Sub

Private Sub cmdNEDMAP_Click()
frmCountryInfo.Hide
frmNEDMAP.Show
End Sub

Private Sub cmdothers_Click()           'msgbox tells user about other countries not included
MsgBox "Andorra, Luxembourg, Monaco, San Marino and The Vatican aren't included as part of 'Western Europe' in this program.  Although isn't that a subjective term anyway?", , "Yeah, about that..."
End Sub

Private Sub cmdPORMAP_Click()
frmCountryInfo.Hide
frmPORMAP.Show
End Sub

Private Sub cmdSUIMAP_Click()
frmCountryInfo.Hide
frmSUIMAP.Show
End Sub

Private Sub cmdUKInfo_Click()                   'again, similar to cmdBELinfo only it is for the UK
picCities.Cls                                   'picture box is cleared to enter in new array (UKcities.txt)
cmdCitiesAsending.Enabled = True
cmdCitiesDesending.Enabled = True
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
UKActive = True                                 'as seen above this activates the sorting buttons to display the corrisponding UKcities.txt arrays
cmdUKtravel.Visible = True
picMap.Picture = LoadPicture(App.Path & "\Images\mapUK.JPG")
picFlag.Picture = LoadPicture(App.Path & "\Images\flagUK.GIF")
    CTR = 0
    Open App.Path & "\UKcities.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
    Next i
cmdGERInfo.Visible = False
cmdIREInfo.Visible = False
cmdESPInfo.Visible = False
cmdFRAInfo.Visible = False
cmdBELInfo.Visible = False
cmdITAInfo.Visible = False
cmdNEDInfo.Visible = False
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True
cmdUKMAP.Visible = True


End Sub
Private Sub cmdGERInfo_Click()                  'the same is repeated for Germany, as above
picCities.Cls
cmdCitiesAsending.Enabled = True
cmdCitiesDesending.Enabled = True
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
GERActive = True
cmdGERtravel.Visible = True
picMap.Picture = LoadPicture(App.Path & "\Images\mapGER.JPG")
picFlag.Picture = LoadPicture(App.Path & "\Images\flagGER.GIF")
 CTR = 0
    Open App.Path & "\GERcities.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
    Next i
cmdGERInfo.Visible = False
cmdIREInfo.Visible = False
cmdESPInfo.Visible = False
cmdFRAInfo.Visible = False
cmdBELInfo.Visible = False
cmdITAInfo.Visible = False
cmdNEDInfo.Visible = False
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True
cmdGERMAP.Visible = True


End Sub
Private Sub cmdNEDInfo_Click()
picCities.Cls
cmdCitiesAsending.Enabled = True
cmdCitiesDesending.Enabled = True
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
NEDActive = True
cmdNEDtravel.Visible = True
picMap.Picture = LoadPicture(App.Path & "\Images\mapNED.JPG")
picFlag.Picture = LoadPicture(App.Path & "\Images\flagNED.GIF")
CTR = 0
    Open App.Path & "\NEDcities.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
    Next i
cmdGERInfo.Visible = False
cmdIREInfo.Visible = False
cmdESPInfo.Visible = False
cmdFRAInfo.Visible = False
cmdBELInfo.Visible = False
cmdITAInfo.Visible = False
cmdNEDInfo.Visible = False
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True
cmdNEDMAP.Visible = True


End Sub
Private Sub cmdFRAInfo_Click()
picCities.Cls
cmdCitiesAsending.Enabled = True
cmdCitiesDesending.Enabled = True
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
FRAActive = True
cmdFRAtravel.Visible = True
picMap.Picture = LoadPicture(App.Path & "\Images\mapFRA.JPG")
picFlag.Picture = LoadPicture(App.Path & "\Images\flagFRA.GIF")
CTR = 0
    Open App.Path & "\FRAcities.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
    Next i
cmdGERInfo.Visible = False
cmdIREInfo.Visible = False
cmdESPInfo.Visible = False
cmdFRAInfo.Visible = False
cmdBELInfo.Visible = False
cmdITAInfo.Visible = False
cmdNEDInfo.Visible = False
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True
cmdFRAMAP.Visible = True


End Sub
Private Sub cmdITAInfo_Click()
picCities.Cls
cmdCitiesAsending.Enabled = True
cmdCitiesDesending.Enabled = True
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
ITAActive = True
cmdITAtravel.Visible = True
picMap.Picture = LoadPicture(App.Path & "\Images\mapITA.JPG")
picFlag.Picture = LoadPicture(App.Path & "\Images\flagITA.GIF")
CTR = 0
    Open App.Path & "\ITAcities.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
    Next i
cmdGERInfo.Visible = False
cmdIREInfo.Visible = False
cmdESPInfo.Visible = False
cmdFRAInfo.Visible = False
cmdBELInfo.Visible = False
cmdITAInfo.Visible = False
cmdNEDInfo.Visible = False
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True
cmdITAMAP.Visible = True


End Sub
Private Sub cmdPORInfo_Click()
picCities.Cls
cmdCitiesAsending.Enabled = True
cmdCitiesDesending.Enabled = True
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
PORActive = True
cmdPORtravel.Visible = True
picMap.Picture = LoadPicture(App.Path & "\Images\mapPOR.JPG")
picFlag.Picture = LoadPicture(App.Path & "\Images\flagPOR.GIF")
CTR = 0
    Open App.Path & "\PORcities.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
    Next i
cmdGERInfo.Visible = False
cmdIREInfo.Visible = False
cmdESPInfo.Visible = False
cmdFRAInfo.Visible = False
cmdBELInfo.Visible = False
cmdITAInfo.Visible = False
cmdNEDInfo.Visible = False
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True
cmdPORMAP.Visible = True


End Sub
Private Sub cmdESPInfo_Click()
picCities.Cls
cmdCitiesAsending.Enabled = True
cmdCitiesDesending.Enabled = True
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
ESPActive = True
cmdESPtravel.Visible = True
picMap.Picture = LoadPicture(App.Path & "\Images\mapESP.JPG")
picFlag.Picture = LoadPicture(App.Path & "\Images\flagESP.GIF")
CTR = 0
    Open App.Path & "\ESPcities.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
    Next i
cmdGERInfo.Visible = False
cmdIREInfo.Visible = False
cmdESPInfo.Visible = False
cmdFRAInfo.Visible = False
cmdBELInfo.Visible = False
cmdITAInfo.Visible = False
cmdNEDInfo.Visible = False
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True
cmdESPMAP.Visible = True


End Sub
Private Sub cmdSUIInfo_Click()
picCities.Cls
cmdCitiesAsending.Enabled = True
cmdCitiesDesending.Enabled = True
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
SUIActive = True
cmdSUItravel.Visible = True
picMap.Picture = LoadPicture(App.Path & "\Images\mapSUI.JPG")
picFlag.Picture = LoadPicture(App.Path & "\Images\flagSUI.GIF")
CTR = 0
    Open App.Path & "\SUIcities.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
    Next i
cmdGERInfo.Visible = False
cmdIREInfo.Visible = False
cmdESPInfo.Visible = False
cmdFRAInfo.Visible = False
cmdBELInfo.Visible = False
cmdITAInfo.Visible = False
cmdNEDInfo.Visible = False
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True
cmdSUIMAP.Visible = True

End Sub
Private Sub cmdIREInfo_Click()
picCities.Cls
cmdCitiesAsending.Enabled = True
cmdCitiesDesending.Enabled = True
cmdABCAsending.Enabled = True
cmdABCDesending.Enabled = True
IREActive = True
cmdIREtravel.Visible = True
picMap.Picture = LoadPicture(App.Path & "\Images\mapIRE.JPG")
picFlag.Picture = LoadPicture(App.Path & "\Images\flagIRE.GIF")
CTR = 0
    Open App.Path & "\IREcities.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Metro(CTR), Population(CTR), Region(CTR)
    Loop
    Close #1
picCities.Print "Metro Area"; Tab(30); "Population"; Tab(50); "Region"
picCities.Print "------------------------------------------------------------------------------------------------------------"
Dim i As Integer
    For i = 1 To CTR
        picCities.Print Metro(i); Tab(30); FormatNumber(Population(i), 0); Tab(50); Region(i)
    Next i
cmdGERInfo.Visible = False
cmdIREInfo.Visible = False
cmdESPInfo.Visible = False
cmdFRAInfo.Visible = False
cmdBELInfo.Visible = False
cmdITAInfo.Visible = False
cmdNEDInfo.Visible = False
cmdPORInfo.Visible = False
cmdUKInfo.Visible = False
cmdSUIInfo.Visible = False
cmdvisible.Visible = True
cmdIREMAP.Visible = True


End Sub

Private Sub cmdUKMAP_Click()                'similar to cmdBELMAP above
frmCountryInfo.Hide
frmUKMAP.Show
End Sub

Private Sub cmdvisible_Click()              'goes back to original list of countries, sets info buttons to visible
cmdGERInfo.Visible = True
cmdESPInfo.Visible = True
cmdFRAInfo.Visible = True
cmdBELInfo.Visible = True
cmdITAInfo.Visible = True
cmdNEDInfo.Visible = True
cmdPORInfo.Visible = True
cmdUKInfo.Visible = True
cmdSUIInfo.Visible = True
cmdIREInfo.Visible = True
cmdvisible.Visible = False                  'sets map buttons to invisible for each case
cmdBELMAP.Visible = False
cmdFRAMAP.Visible = False
cmdGERMAP.Visible = False
cmdIREMAP.Visible = False
cmdITAMAP.Visible = False
cmdUKMAP.Visible = False
cmdNEDMAP.Visible = False
cmdPORMAP.Visible = False
cmdSUIMAP.Visible = False
cmdESPMAP.Visible = False

cmdBELtravel.Visible = False                   'sets website link buttons to invisible for each case
cmdFRAtravel.Visible = False
cmdGERtravel.Visible = False
cmdIREtravel.Visible = False
cmdITAtravel.Visible = False
cmdUKtravel.Visible = False
cmdNEDtravel.Visible = False
cmdPORtravel.Visible = False
cmdSUItravel.Visible = False
cmdESPtravel.Visible = False
End Sub

Private Sub cmdBELtravel_Click()            'this is made visible when cmdBELinfo is hit
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/Belgium", "", "", True      'links to wikitravel page on Belgium

End Sub
Private Sub cmdESPtravel_Click()            'same as above but for Spain
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/Spain", "", "", True
End Sub
Private Sub cmdFRAtravel_Click()
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/France", "", "", True
End Sub
Private Sub cmdGERtravel_Click()
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/Germany", "", "", True
End Sub
Private Sub cmdIREtravel_Click()
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/Ireland", "", "", True
End Sub
Private Sub cmdITAtravel_Click()
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/Italy", "", "", True
End Sub
Private Sub cmdNEDtravel_Click()
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/Netherlands", "", "", True
End Sub
Private Sub cmdPORtravel_Click()
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/Portugal", "", "", True
End Sub
Private Sub cmdSUItravel_Click()
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/Switzerland", "", "", True
End Sub
Private Sub cmdUKtravel_Click()
    ShellExecute Me.hWnd, "open", "http://wikitravel.org/en/United_Kingdom", "", "", True
End Sub
Private Sub Form_Load()         'when the form is loaded the detail map buttons are set to invisible, gets rid of clutter
picCities.Cls
cmdvisible.Visible = False
cmdBELMAP.Visible = False
cmdFRAMAP.Visible = False
cmdGERMAP.Visible = False
cmdIREMAP.Visible = False
cmdITAMAP.Visible = False
cmdUKMAP.Visible = False
cmdNEDMAP.Visible = False
cmdPORMAP.Visible = False
cmdSUIMAP.Visible = False
cmdESPMAP.Visible = False
picMap.Picture = LoadPicture(App.Path & "\Images\europemap.GIF")        'default europe map is displayed

cmdUKtravel.Visible = False     'travel buttons are also made invisible
cmdBELtravel.Visible = False
cmdFRAtravel.Visible = False
cmdGERtravel.Visible = False
cmdIREtravel.Visible = False
cmdITAtravel.Visible = False
cmdUKtravel.Visible = False
cmdNEDtravel.Visible = False
cmdPORtravel.Visible = False
cmdSUItravel.Visible = False
End Sub

