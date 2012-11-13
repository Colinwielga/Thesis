VERSION 5.00
Begin VB.Form frmCurrency 
   BackColor       =   &H00000000&
   Caption         =   "Currency Converter"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000A&
      Height          =   2175
      Left            =   3720
      ScaleHeight     =   2115
      ScaleWidth      =   3315
      TabIndex        =   12
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtDesired 
      BackColor       =   &H8000000A&
      Height          =   735
      Left            =   3720
      TabIndex        =   11
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtOriginal 
      BackColor       =   &H8000000B&
      Height          =   735
      Left            =   3720
      TabIndex        =   10
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtAmount 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H0000C000&
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000C000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton cmdReturn3 
      BackColor       =   &H0000C000&
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Lblrates 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "* When traveling make sure to pay attention to transaction rates, as they do vary from location to location!"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2205
      Left            =   1320
      Picture         =   "frmCurrency.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   2250
   End
   Begin VB.Label lblConverted 
      BackColor       =   &H80000012&
      Caption         =   "Converted Currency"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   7200
      TabIndex        =   16
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblDesired 
      BackColor       =   &H00000000&
      Caption         =   "Desired Currency (1-6)"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   7200
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblOriginal 
      BackColor       =   &H80000012&
      Caption         =   "Original Currency (1-6)"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   7200
      TabIndex        =   14
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00000000&
      Caption         =   "Amount to be converted"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   7200
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblNZ 
      BackColor       =   &H80000012&
      Caption         =   "6. Canadian Dollar"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lblAus 
      BackColor       =   &H80000012&
      Caption         =   "5. Australian Dollar"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label lblYen 
      BackColor       =   &H80000012&
      Caption         =   "4. Japanese Yen"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblEuro 
      BackColor       =   &H80000012&
      Caption         =   "3. Euro"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblGBP 
      BackColor       =   &H80000012&
      Caption         =   "2. British Pound"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label LblUS 
      BackColor       =   &H00000000&
      Caption         =   "1. U.S. Dollar"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This application converts a user given currency of six different choices to another.
'The user inputs a monetary value into the first text box, then inputs a 1-6 into the
'second text box, this indicates the currency of their first input.  Finally the user
'inputs another value between 1-6, this indicates the desired currency to conver to.

Dim ctr As Integer      'Sets the varaibles of the application
Dim Amount As Single
Dim Original As Integer
Dim Desired As Integer
Dim ConversionFactors(1 To 6) As Single
Dim sum As Single

Private Sub cmdConvert_Click()

Open App.Path & "\ConversionData.txt" For Input As #1       'Opens the file that contains the currency information

Amount = txtAmount.Text     'Initiates variables
ctr = 0
Original = txtOriginal.Text
Desired = txtDesired.Text

Do While Not EOF(1)     'Loops through file until the end
    ctr = ctr + 1       'Adds one to the counter
    Input #1, ConversionFactors(ctr)        'Inputs the file data
Loop
    sum = Amount * (ConversionFactors(Original) * ConversionFactors(Desired))    'Calculates the converted currency
    picResults.Print FormatCurrency(sum, 2)     'Displays the new converted currency in the picture box

Close #1        'Closes the file
End Sub

Private Sub cmdClear_Click()
 picResults.Cls                 'Clears all text boxes and picture box
    txtAmount.Text = ""
    txtOriginal.Text = ""
    txtDesired.Text = ""
End Sub

Private Sub cmdReturn3_Click()
frmCurrency.Hide                'retruns to main form
FrmMain.Show                    'hides this form
End Sub


