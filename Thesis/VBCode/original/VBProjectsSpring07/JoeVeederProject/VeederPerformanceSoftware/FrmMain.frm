VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FF0000&
   Caption         =   "Car Performance Software"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicVirtualRace 
      Height          =   2535
      Left            =   3600
      ScaleHeight     =   2475
      ScaleWidth      =   3555
      TabIndex        =   48
      Top             =   7440
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton CmdSkyline 
      BackColor       =   &H000000C0&
      Height          =   735
      Left            =   840
      Picture         =   "FrmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Nissan Skyline"
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton CmdFord 
      BackColor       =   &H000000C0&
      Height          =   735
      Left            =   840
      Picture         =   "FrmMain.frx":6706
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Ford F-150"
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton CmdDragster 
      BackColor       =   &H000000C0&
      Height          =   735
      Left            =   840
      Picture         =   "FrmMain.frx":DF64
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Top Fuel Dragster"
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton CmdMiata 
      BackColor       =   &H000000C0&
      Height          =   735
      Left            =   840
      Picture         =   "FrmMain.frx":14B5D
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Mazda Miata"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.PictureBox PicNitrousCheck 
      Height          =   495
      Left            =   8640
      Picture         =   "FrmMain.frx":1BCD6
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   43
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox PicTurboCheck 
      Height          =   495
      Left            =   8640
      Picture         =   "FrmMain.frx":1C17E
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   42
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdNitrous 
      BackColor       =   &H000000C0&
      Caption         =   "Add A Nitrous Oxide System"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox TxtHPMirror 
      Height          =   375
      Left            =   6960
      TabIndex        =   39
      Text            =   "0"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton CmdTurbo 
      BackColor       =   &H000000C0&
      Caption         =   "Add A Turbocharger"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox TxtNewHP 
      Height          =   375
      Left            =   6960
      TabIndex        =   35
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton CmdVirtualRace 
      BackColor       =   &H000000C0&
      Caption         =   "Virtual Race"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton CmdCredits 
      BackColor       =   &H000000C0&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H000000C0&
      Caption         =   "Clear Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton CmdCalculate 
      BackColor       =   &H000000C0&
      Caption         =   "Calculate Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton CmdKilograms 
      BackColor       =   &H000000C0&
      Caption         =   "Convert From Kilograms"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton CmdWatt 
      BackColor       =   &H000000C0&
      Caption         =   "Convert From Kilowatts"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1560
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   2520
      Picture         =   "FrmMain.frx":1C626
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   26
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox TxtWeight 
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Text            =   "0"
      ToolTipText     =   "Enter Your Vehicles Weight in Pounds"
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox TxtHP 
      Height          =   375
      Left            =   960
      TabIndex        =   22
      Text            =   "0"
      ToolTipText     =   "Enter Your Vehicles Horsepower"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TxtTire 
      Height          =   375
      Left            =   13680
      TabIndex        =   20
      Text            =   "0"
      ToolTipText     =   "Enter A Tire Size Between 13 and 18 inches"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   11760
      Picture         =   "FrmMain.frx":23DAC
      ScaleHeight     =   1635
      ScaleWidth      =   1635
      TabIndex        =   18
      Top             =   6000
      Width           =   1695
   End
   Begin VB.PictureBox PicMPH 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   9360
      Width           =   1455
   End
   Begin VB.PictureBox PicET 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   1395
      TabIndex        =   10
      Top             =   8040
      Width           =   1455
   End
   Begin VB.PictureBox PicTruck 
      Height          =   615
      Left            =   12840
      Picture         =   "FrmMain.frx":2CEEA
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.PictureBox PicSUV 
      Height          =   615
      Left            =   12840
      Picture         =   "FrmMain.frx":2D253
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.PictureBox PicTwoDoor 
      FillColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   12840
      Picture         =   "FrmMain.frx":2D696
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox PicFourDoor 
      Height          =   615
      Left            =   12840
      Picture         =   "FrmMain.frx":2D9ED
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.PictureBox PicConvertible 
      Height          =   615
      Left            =   12840
      Picture         =   "FrmMain.frx":2DD57
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.PictureBox PicMain 
      Height          =   3255
      Left            =   4440
      Picture         =   "FrmMain.frx":2E071
      ScaleHeight     =   3195
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   3120
      Width           =   6015
      Begin VB.Line Line1 
         X1              =   3120
         X2              =   3000
         Y1              =   3240
         Y2              =   3240
      End
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H000000C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   1335
   End
   Begin VB.PictureBox PicTwoDoorHighlight 
      BackColor       =   &H0000C000&
      Height          =   855
      Left            =   12720
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicFourDoorHighlight 
      BackColor       =   &H0000C000&
      Height          =   855
      Left            =   12720
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicConvertibleHighlight 
      BackColor       =   &H0000C000&
      Height          =   855
      Left            =   12720
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicSUVHighlight 
      BackColor       =   &H0000C000&
      Height          =   855
      Left            =   12720
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicTruckHighlight 
      BackColor       =   &H0000C000&
      Height          =   855
      Left            =   12720
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicDisplayPreset 
      Height          =   10440
      Left            =   -120
      Picture         =   "FrmMain.frx":402EB
      ScaleHeight     =   10380
      ScaleWidth      =   16515
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   16575
      Begin VB.Label LblMPH 
         Caption         =   "Miles Per Hour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   63
         Top             =   8760
         Width           =   1935
      End
      Begin VB.Label LblET 
         Caption         =   "Quarter Mile Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   62
         Top             =   7440
         Width           =   2295
      End
      Begin VB.Label LblInches 
         Caption         =   "(13-18) Inches"
         Height          =   255
         Left            =   13800
         TabIndex        =   61
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Tire Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13800
         TabIndex        =   60
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Aerodynamics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   59
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "    New Engine Horsepower"
         Height          =   255
         Left            =   6480
         TabIndex        =   58
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "   Current Engine Horsepower"
         Height          =   255
         Left            =   6480
         TabIndex        =   57
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "  Engine Modifications"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   56
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label LblPresets 
         Caption         =   "  Preset Cars"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   55
         Top             =   6720
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "  Enter Vehicle Weight (lbs)"
         Height          =   255
         Left            =   720
         TabIndex        =   54
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label LblVehicleWeight 
         Caption         =   " Vehicle Weight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   53
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label LblEnterHp 
         Caption         =   "  Enter Horsepower"
         Height          =   255
         Left            =   840
         TabIndex        =   52
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label LblEngine 
         Caption         =   "   Engine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   50
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "   Engine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   51
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label LblHP 
      BackColor       =   &H0000C000&
      Caption         =   "  Enter Horsepower"
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label LblCurrentHP 
      BackColor       =   &H0000C000&
      Caption         =   "   Current Engine Horsepower"
      Height          =   255
      Left            =   6360
      TabIndex        =   40
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label LblNewHP 
      BackColor       =   &H0000C000&
      Caption         =   "      New Engine Horsepower"
      Height          =   255
      Left            =   6360
      TabIndex        =   38
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label LblEngineModifications 
      BackColor       =   &H0000C000&
      Caption         =   "  Engine Modifications"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   36
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label LblPresetCarValues 
      BackColor       =   &H0000C000&
      Caption         =   "  Preset Cars"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   34
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label LblType 
      BackColor       =   &H0000C000&
      Caption         =   "  Click To Choose Vehicle Type"
      Height          =   255
      Left            =   12240
      TabIndex        =   30
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label LblEnterWeight 
      BackColor       =   &H0000C000&
      Caption         =   "  Enter Vehicle Weight (lbs)"
      Height          =   255
      Left            =   600
      TabIndex        =   24
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label LblWeight 
      BackColor       =   &H0000C000&
      Caption         =   " Vehicle Weight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label LblEnterTireSize 
      BackColor       =   &H0000C000&
      Caption         =   " 13-17 Inches"
      Height          =   255
      Left            =   13680
      TabIndex        =   19
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label LblTireSize 
      BackColor       =   &H0000C000&
      Caption         =   " Tire Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   17
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label LblMilesPerHour 
      BackColor       =   &H000080FF&
      Caption         =   "   Miles Per Hour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   9
      Top             =   8640
      Width           =   2295
   End
   Begin VB.Label LblQuarterMileTime 
      BackColor       =   &H000080FF&
      Caption         =   "Quarter Mile Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label LblAerodynamics 
      BackColor       =   &H0000C000&
      Caption         =   " Aerodynamics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'checks all variables
Option Explicit
'defines the form level variables
Dim MPH As Single
Dim ET As Single
Dim HP As Single
Dim AERO As Single
'This subroutine calculates the quarter mile time and mph based on user inputs
Private Sub CmdCalculate_Click()
    'defines the variables used in the calculate function
    Dim WEIGHT As Single
    Dim TIRE As Single
    Dim DRAG As Single
    'gets the variables from user input into the text boxes
    HP = TxtHP.Text
    WEIGHT = TxtWeight.Text
    TIRE = TxtTire.Text
    TxtHPMirror.Text = TxtHP.Text
    'displays an error message if no user input is made for horsepower
    If HP = 0 Then
        MsgBox ("Please Input Your Vehicles Horsepower")
    End If
    'displays an error message if the user inputs a negative number for horsepower
    If HP < 0 Then
        MsgBox ("Please Enter a Positive Number for Horsepower")
    End If
    'displays an error message if the user does not click an aerodynamic type
    If AERO = 0 Then
        MsgBox ("Please Select A Vehicle Type")
    End If
    'displays an error message if the user does not input a weight
    If WEIGHT = 0 Then
        MsgBox ("Please Enter Your Vehicle Weight")
    End If
    'displays an error message if the user inputs a negative number for weight
    If WEIGHT < 0 Then
        MsgBox ("Please Enter a Positive Number for Weight")
    End If
    'adjusts the horsepower in order to compensate for tire size
    Select Case TIRE
        Case Is = 13
            HP = HP * 0.956
        Case Is = 14
            HP = HP
        Case Is = 15
            HP = HP * 1.044
        Case Is = 16
            HP = HP * 1.088
        Case Is = 17
            HP = HP * 1.132
        Case Is = 18
            HP = HP * 1.176
        'displays an error message if no tire size is chosen
        Case Is = 0
            MsgBox ("Please Select A Tire Size")
        'displays an error message if any other variable is entered
        Case Else
            MsgBox ("Please Enter A Tire Size Between 13 and 18 Inches")
    End Select
    'if the user completes all inputs, the mph is calculated using a formula
    If HP > 0 And WEIGHT > 0 Then
        MPH = ((HP / WEIGHT) ^ 0.3003) * 213.83
    End If
    'deducts horsepower based on aerodynamic drag, aero is the variable saved with the drag coefficient for each vehicle type
    DRAG = 0.005 * AERO * 20 * (MPH ^ 3) * 0.00134
    HP = HP - DRAG
    'if all inputs are completed, mph is calulated using the hp taking into account aerodynamic drag
    If HP > 0 And WEIGHT > 0 Then
        MPH = ((HP / WEIGHT) ^ 0.3003) * 213.83
        'quarter mile time is calculated using the revised hp
        ET = ((WEIGHT / HP) ^ 0.2545) * 7.4499
    End If
    'if each input is completed, the pic box is cleared and et and mph displayed
    If TIRE <> 0 And WEIGHT <> 0 And AERO <> 0 And HP <> 0 Then
        PicET.Cls
        PicMPH.Cls
        'displays quarter mile time to two places
        PicET.Print FormatNumber(ET, 2)
        'displays miles per hour to two places
        PicMPH.Print FormatNumber(MPH, 2)
    End If
    End Sub
'clear button sets every input to zero
Private Sub CmdClear_Click()
    TxtHP.Text = 0
    TxtWeight.Text = 0
    TxtTire.Text = 0
    TxtHPMirror.Text = 0
    TxtNewHP.Text = 0
    'clears the display boxes
    PicET.Cls
    PicMPH.Cls
    'sets the background back to its original settings
    PicDisplayPreset.Visible = False
    PicMain.Visible = True
    'clears the virtual race box
    PicVirtualRace.Cls
    'sets the check marks to not visible
    PicVirtualRace.Visible = False
    PicNitrousCheck.Visible = False
    PicTurboCheck.Visible = False
    ET = 0
End Sub
'displays the form that lists the program credits
Private Sub CmdCredits_Click()
    FrmMain.Visible = False
    FrmCredits.Visible = True
End Sub
'subroutine to print a virtual listing of how the user input vehicle compares to 9 other vehicles in a race
Private Sub CmdVirtualRace_Click()
    'defines the variables
    Dim Cars(1 To 10) As String
    Dim Times(1 To 10) As Single
    Dim Ctr As Single
    Dim Place As Integer
    Dim Printed As Integer
    'if an answer is present for the quarter mile time, a virtual ranking is shown
    If ET > 0 Then
        'makes the output box visible
        PicVirtualRace.Visible = True
        PicVirtualRace.Cls
        'prints the title
        PicVirtualRace.Print "                  Virtual Race Standings"
        PicVirtualRace.Print ""
        'opens a text file with each of the 9 preset vehicles
        Open App.Path & "\Times.txt" For Input As #1
        Ctr = 0
        'places the cars and times into an array
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Cars(Ctr), Times(Ctr)
        Loop
        'closes the text file
        Close #1
        Printed = 1
        'prints each car and time
        For Place = 1 To Ctr
            PicVirtualRace.Print Cars(Place); Tab(30); Times(Place); Tab(38); "(Sec)"
            If ET < Times(Place + 1) And Printed = 1 Then
                PicVirtualRace.Print "User Inputed Vehicle"; Tab(31); FormatNumber(ET, 2); Tab(38); "(Sec)"
                'sets printed to two so the user vehicle is only printed once
                Printed = 2
            End If
        Next Place
        If Printed = 1 Then
            'prints the user input time if it was slower then each of the others, and not yet printed
            PicVirtualRace.Print "User Inputed Vehicle"; Tab(31); FormatNumber(ET, 2); Tab(38); "(Sec)"
        End If
    'prints error if all fields are not completed
    Else: MsgBox ("Please Complete All Fields And Click Calculate Before Using This Feature")
    End If
End Sub
'changes the horsepower value if the nitrous box is clicked
Private Sub CmdNitrous_Click()
    'makes a mirror image of the other HP text box
    TxtHPMirror.Text = TxtHP.Text
    'makes the nitrous check mark visible when clicked
    PicNitrousCheck.Visible = True
    'adds 50 horsepower to HP if an HP value has been input
    If TxtHP.Text > 0 Then
        TxtHP.Text = TxtHP.Text + 50
    'displays an error if no HP has been input
    Else: MsgBox ("Please Enter Your Horsepower")
        PicNitrousCheck.Visible = False
    End If
    'places the HP value in the results HP box
    TxtNewHP.Text = TxtHP.Text
End Sub
'changes the HP value if the turbo box is clicked
Private Sub CmdTurbo_Click()
    'makes a mirror image of the other HP text box
    TxtHPMirror.Text = TxtHP.Text
    'makes the turbo check mark visible when clicked
    PicTurboCheck.Visible = True
    'if the HP has been input, is is increased by 35%
    If TxtHP.Text > 0 Then
        TxtHP.Text = TxtHP.Text * 1.35
    'displays an error if no HP has been input
    Else: MsgBox ("Please Enter Your Horsepower")
        PicTurboCheck.Visible = False
    End If
    'places the HP value in the results HP box
    TxtNewHP.Text = TxtHP.Text
End Sub
'changes the aerodynamic value to .001 for a Two Door Car
Private Sub PicTwoDoor_Click()
    'highlights only the Two Door car box
    PicTwoDoorHighlight.Visible = True
    PicFourDoorHighlight.Visible = False
    PicConvertibleHighlight.Visible = False
    PicSUVHighlight.Visible = False
    PicTruckHighlight.Visible = False
    AERO = 0.001
End Sub
'changes the aerodynamic value to .01 for a Four Door Car
Private Sub PicFourDoor_Click()
    'highlights only the Four Door car box
    PicTwoDoorHighlight.Visible = False
    PicFourDoorHighlight.Visible = True
    PicConvertibleHighlight.Visible = False
    PicSUVHighlight.Visible = False
    PicTruckHighlight.Visible = False
    AERO = 0.01
End Sub
'changes the aerodynamic value to .06 for a Convertible
Private Sub PicConvertible_Click()
    'highlights only the Convertible car box
    PicTwoDoorHighlight.Visible = False
    PicFourDoorHighlight.Visible = False
    PicConvertibleHighlight.Visible = True
    PicSUVHighlight.Visible = False
    PicTruckHighlight.Visible = False
    AERO = 0.06
End Sub
'changes the aerodynamic value to .07 for an SUV
Private Sub PicSUV_Click()
    'highlights only the SUV car box
    PicTwoDoorHighlight.Visible = False
    PicFourDoorHighlight.Visible = False
    PicConvertibleHighlight.Visible = False
    PicSUVHighlight.Visible = True
    PicTruckHighlight.Visible = False
    AERO = 0.07
End Sub
'changes the aerodynamic value to .28 for a Pickup
Private Sub PicTruck_Click()
    'highlights only the Pickup car box
    PicTwoDoorHighlight.Visible = False
    PicFourDoorHighlight.Visible = False
    PicConvertibleHighlight.Visible = False
    PicSUVHighlight.Visible = False
    PicTruckHighlight.Visible = True
    AERO = 0.28
End Sub
'changes the weight box input into pounds
Private Sub CmdKilograms_Click()
    TxtWeight.Text = TxtWeight.Text * 2.20462262
End Sub
'changes the Horsepower box input into HP from Watts
Private Sub CmdWatt_Click()
    TxtHP.Text = TxtHP.Text * 1.34102209
End Sub
'inputs preset values into each of the text boxes for a mazda miata
Private Sub CmdMiata_Click()
    AERO = 0.06
    PicTurboCheck.Visible = False
    PicNitrousCheck.Visible = False
    TxtNewHP.Text = 0
    TxtHP.Text = 140
    TxtHPMirror.Text = 0
    TxtWeight.Text = 2332
    TxtTire.Text = 15
    PicTwoDoorHighlight.Visible = True
    PicFourDoorHighlight.Visible = False
    PicConvertibleHighlight.Visible = False
    PicSUVHighlight.Visible = False
    PicTruckHighlight.Visible = False
End Sub
'inputs preset values into each of the text boxes for a Top Fuel Dragster
Private Sub CmdDragster_Click()
    AERO = 0.001
    PicTurboCheck.Visible = True
    PicNitrousCheck.Visible = False
    TxtNewHP.Text = 0
    TxtHP.Text = 7000
    TxtHPMirror.Text = 0
    TxtWeight.Text = 2225
    TxtTire.Text = 18
    PicTwoDoorHighlight.Visible = True
    PicFourDoorHighlight.Visible = False
    PicConvertibleHighlight.Visible = False
    PicSUVHighlight.Visible = False
    PicTruckHighlight.Visible = False
End Sub
'inputs preset values into each of the text boxes for a Nissan Skyline
Private Sub CmdSkyline_Click()
    AERO = 0.001
    PicTurboCheck.Visible = True
    PicNitrousCheck.Visible = False
    TxtNewHP.Text = 0
    TxtHP.Text = 330
    TxtHPMirror.Text = 0
    TxtWeight.Text = 3672
    TxtTire.Text = 18
    PicTwoDoorHighlight.Visible = True
    PicFourDoorHighlight.Visible = False
    PicConvertibleHighlight.Visible = False
    PicSUVHighlight.Visible = False
    PicTruckHighlight.Visible = False
    PicDisplayPreset.Visible = True
    PicMain.Visible = False
End Sub
'inputs preset values into each of the text boxes for a Ford F-150
Private Sub CmdFord_Click()
    AERO = 0.28
    PicTurboCheck.Visible = False
    PicNitrousCheck.Visible = False
    TxtNewHP.Text = 0
    TxtHP.Text = 230
    TxtHPMirror.Text = 0
    TxtWeight.Text = 4690
    TxtTire.Text = 17
    PicTwoDoorHighlight.Visible = False
    PicFourDoorHighlight.Visible = False
    PicConvertibleHighlight.Visible = False
    PicSUVHighlight.Visible = False
    PicTruckHighlight.Visible = True
End Sub
'Ends the program
Private Sub CmdQuit_Click()
    End
End Sub

