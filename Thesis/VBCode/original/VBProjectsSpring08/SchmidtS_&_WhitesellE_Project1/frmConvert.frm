VERSION 5.00
Begin VB.Form frmConvert 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Currency Converter"
   ClientHeight    =   6795
   ClientLeft      =   2730
   ClientTop       =   2325
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10260
   Begin VB.TextBox txtOriginal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   15
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Return to Programs Page"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Click to Enter Length to Convert"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtDesired 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   9
      Top             =   4200
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   6120
      Picture         =   "frmConvert.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "7. U.S. Dollar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   14
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblRand 
      BackColor       =   &H00C0FFC0&
      Caption         =   "8.  South African Rand"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   13
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblDesired 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Desired Currency (1-8):"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label lblOriginal 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Original Currency (1-8): "
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label lblPound 
      BackColor       =   &H00C0FFC0&
      Caption         =   "6.  British Pound"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   6
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblChilPeso 
      BackColor       =   &H00C0FFC0&
      Caption         =   "5.  Chilean Peso"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblYuan 
      BackColor       =   &H00C0FFC0&
      Caption         =   "4.  Yuan (China)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblAustDol 
      BackColor       =   &H00C0FFC0&
      Caption         =   "3.  Australian Dollar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblYen 
      BackColor       =   &H00C0FFC0&
      Caption         =   "2.  Yen (Japan)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblEuro 
      BackColor       =   &H00C0FFC0&
      Caption         =   "1.  Euro"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Currency Converter"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   10335
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this page will take convert a currency to a different currency
'written 3/26/08 by Sammi and Erika


Private Sub cmdConvert_Click()
Dim Amount As Single, Original As Single, Desired As Single
Dim Conversion(1 To 8) As Single
Dim ConvertedValue As Single


'gets amount to be converted from user using an input box and the numbers (which correspond to different currencies) of the
'original currency and desired currency

'1 euro = 1.5779 dollars
'1 yen = 0.01 dollars
'1 australian dollar = 0.9187 dollars
'1 yuan = 0.1425 dollars
'1 chilean peso = 0.002 dollars
'1 british pound = 2.007 dollars
'1 dollar = 1 dollar
'1 south african rand = 0.1251


Amount = InputBox("Enter the amount to be converted, then push enter.", "Amount To Convert")
Original = txtOriginal.Text
Desired = txtDesired.Text

CTR = 0

'opens file with conversion rates and reads the list into an array

Open App.Path & "\conversion.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Conversion(CTR)
    Loop
Close #1

ConvertedValue = Amount * Conversion(Original) / Conversion(Desired)

MsgBox "The converted value is " & FormatNumber(ConvertedValue, 2) & ".", , "Converted Currency"
    

End Sub

Private Sub cmdGoBack_Click()
frmConvert.Hide
frmPrograms.Show
End Sub

