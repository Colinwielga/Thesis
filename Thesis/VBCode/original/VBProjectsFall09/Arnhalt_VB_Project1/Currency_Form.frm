VERSION 5.00
Begin VB.Form frmCurrency 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Convert Currency"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3878
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdGoToHome 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to the Home Page"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1598
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txtBegAmount 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4530
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CONVERT"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2370
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2535
   End
   Begin VB.OptionButton optDollarsToPounds 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Convert DOLLARS to POUNDS"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4050
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.OptionButton optPoundsToDollars 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Convert POUNDS to DOLLARS"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1058
      TabIndex        =   0
      Top             =   1080
      Width           =   2520
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   758
      Picture         =   "Currency_Form.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   600
      Width           =   6015
      Begin VB.Label lblBeginningAmount 
         BackColor       =   &H00FF8080&
         Caption         =   "Enter your starting amount -->"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   2895
      End
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: The currency is converted using exchange rates as of October 19, 2009."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1050
      TabIndex        =   6
      Top             =   4680
      Width           =   5655
   End
End
Attribute VB_Name = "frmCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: London Attractions
'Form Name: Convert Currency
'Author: Heather Arnhalt
'Date Written: October 18, 2009
'Objective: The user chooses whether they would like to convert dollars to pounds or pounds to dollars using option boxes.
'Depending upon which option they choose, the program uses the conversion rate as of October 19 (found on yahoo.com) to calculate
'the correct conversion amount.

Private Sub cmdCalculate_Click()

    'declare the values used for this subroutine
    Dim startingValue As String, convertedValue As String

    'get the starting value what the user entered in the text box
    startingValue = txtBegAmount.Text

    'calculate pounds to dollars if the pounds to dollars option box was selected or
    'dollars to pounds if the dollars to pounds option box was selected by multiplying the starting value by the appropriate conversion rate
    If optPoundsToDollars.Value = True Then
        convertedValue = startingValue * 1.6312
        MsgBox FormatNumber(startingValue, 2) & " British pounds is equal to " & FormatCurrency(convertedValue) & ".", , "Conversion"
    Else
        convertedValue = startingValue * 0.6131
        MsgBox FormatCurrency(startingValue) & " is equal to " & FormatNumber(convertedValue, 2) & " British pounds.", , "Conversion"
    End If

End Sub

Private Sub cmdGoToHome_Click()
    'hide the Currency form and show the Home Page form
    frmCurrency.Hide
    frmHomePage.Show
End Sub

Private Sub cmdQuit_Click()
    'end the program
    End
End Sub
