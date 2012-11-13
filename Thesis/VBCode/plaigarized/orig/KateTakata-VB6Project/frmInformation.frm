VERSION 5.00
Begin VB.Form frmInformation 
   BackColor       =   &H000080FF&
   Caption         =   "Before You Begin..."
   ClientHeight    =   9975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15930
   FillColor       =   &H80000001&
   BeginProperty Font 
      Name            =   "Segoe UI Symbol"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   15930
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   9720
      ScaleHeight     =   4335
      ScaleWidth      =   5895
      TabIndex        =   12
      Top             =   2760
      Width           =   5895
   End
   Begin VB.TextBox txtAllInfo 
      Height          =   6375
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2280
      Width           =   6855
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      TabIndex        =   5
      Top             =   8880
      Width           =   3975
   End
   Begin VB.CommandButton cmdGel 
      Caption         =   "Information on Gels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdBeginCat 
      Caption         =   "Let's work on the Catwalk! "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   3
      Top             =   8880
      Width           =   4095
   End
   Begin VB.CommandButton cmdEquipment 
      Caption         =   "Information on Lighting Instruments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdElectrics 
      Caption         =   "Information on Electrics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "What are we doing?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Here's some information to point you in the right direction. Click on the buttons below to learn more about what you'll be doing!"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   12375
   End
   Begin VB.Label lblColors 
      BackStyle       =   0  'Transparent
      Caption         =   "The gels you'll have available: R09, R25, R62, R80, R89 (L-R)"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   16
      Top             =   7200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblEquip 
      BackStyle       =   0  'Transparent
      Caption         =   "The lighting equipment you'll be using: PAR, Ellipsoidal, Scoop, Fresnel (clockwise left corner)"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   15
      Top             =   7200
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label lblCatPic 
      BackStyle       =   0  'Transparent
      Caption         =   "An electric in the Gorecki Theater"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   14
      Top             =   7200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblPlotExample 
      BackStyle       =   0  'Transparent
      Caption         =   "An example of a light plot"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11520
      TabIndex        =   13
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblEquipment 
      BackStyle       =   0  'Transparent
      Caption         =   "Lighting Equipment"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblGels 
      BackStyle       =   0  'Transparent
      Caption         =   "Gels"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblElectric 
      BackStyle       =   0  'Transparent
      Caption         =   "Electrics"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblPlot 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "What the Heck is a Light Plot?!"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   14775
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is an information display form.

Private Sub cmdPlot_Click()
'Displays information and a corresponding picture about what the program will do.

    Dim NextLine As String
    Dim Message As String
    
    lblPlot.Visible = True
    lblElectric.Visible = False
    lblEquipment.Visible = False
    lblGels.Visible = False
    
    Open App.Path & "\TextPlot.txt" For Input As #1
    Message = "In this program, you will be creating a light plot."
    
    Do Until EOF(1)
        Input #1, NextLine
        Message = Message & NextLine & " "
    Loop
    
    Close #1
    
    txtAllInfo.Text = Message
    picResults.Picture = LoadPicture(App.Path & "\110_liteplot - PAINT.JPG")
    
    lblPlotExample.Visible = True
    lblCatPic.Visible = False
    lblColors.Visible = False
    lblEquip.Visible = False
    
End Sub

Private Sub cmdElectrics_Click()
'Displays information and a corresponding picture about light electrics.

    Dim NextLine As String
    Dim Message As String
    
    lblElectric.Visible = True
    lblPlot.Visible = False
    lblEquipment.Visible = False
    lblGels.Visible = False
    
    Open App.Path & "\TextElectrics.txt" For Input As #1
    Message = "Electrics are where the lights are hung and plugged into."
    
    Do Until EOF(1)
        Input #1, NextLine
        Message = Message & NextLine & " "
    Loop
    
    Close #1
    
    txtAllInfo.Text = Message
    picResults.Picture = LoadPicture(App.Path & "\Electric.JPG")
    
    lblCatPic.Visible = True
    lblPlotExample.Visible = False
    lblColors.Visible = False
    lblEquip.Visible = False
    
End Sub

Private Sub cmdEquipment_Click()
'Displays information and a corresponding picture about lighting instruments.

    Dim NextLine As String
    Dim Message As String
    
    lblPlot.Visible = False
    lblElectric.Visible = False
    lblEquipment.Visible = True
    lblGels.Visible = False
    
    Open App.Path & "\TextInstruments.txt" For Input As #1
    Message = "Different types of lights have different functions and are placed best in certain spots."
    
    Do Until EOF(1)
        Input #1, NextLine
        Message = Message & NextLine & " "
    Loop
    
    Close #1
    
    txtAllInfo.Text = Message
    
    picResults.Picture = LoadPicture(App.Path & "\4Lights.JPG")
    
    lblEquip.Visible = True
    lblColors.Visible = False
    lblPlotExample.Visible = False
    lblCatPic.Visible = False
    
End Sub

Private Sub cmdGel_Click()
'Displays information and a corresponding picture about gels.

    Dim NextLine As String
    Dim Message As String
    
    lblPlot.Visible = False
    lblElectric.Visible = False
    lblEquipment.Visible = False
    lblGels.Visible = True
    
    Open App.Path & "\TextGels.txt" For Input As #1
    Message = "Gels are sheets of thin colored plastic."
    
    Do Until EOF(1)
        Input #1, NextLine
        Message = Message & NextLine & " "
    Loop
    
    Close #1
    
    txtAllInfo.Text = Message
    picResults.Picture = LoadPicture(App.Path & "\5Colors.JPG")
    
    lblColors.Visible = True
    lblPlotExample.Visible = False
    lblCatPic.Visible = False
    lblEquip.Visible = False
    
End Sub

Private Sub cmdBeginCat_Click()
'Advances the program to the Catwalk form.
    frmInformation.Hide
    frmCatwalk.Show
End Sub

Private Sub cmdQuit_Click()
'Ends the program.
    End
End Sub




