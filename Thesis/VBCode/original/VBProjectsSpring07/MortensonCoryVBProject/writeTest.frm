VERSION 5.00
Begin VB.Form frmSave 
   BackColor       =   &H00000000&
   Caption         =   "Save Results"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWithheld 
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox txtIncome 
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox txtLiability 
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtAGI 
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save "
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00000000&
      Caption         =   "Taxes Withheld"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1560
      TabIndex        =   10
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00000000&
      Caption         =   "Tax Liability"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00000000&
      Caption         =   "Taxable Income:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00000000&
      Caption         =   "Adjusted Gross Income:"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00000000&
      Caption         =   "Enter And Save The Following Information:"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   8415
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSave_Click()
'This operation will write all of the user's information obtained into a data file to be manipulated at a different date.
    Dim AGIInput As Double, Income As Double, Liability As Double
    Dim Withheld As Double
    
    AGIInput = txtAGI.Text
    Income = txtIncome.Text
    Liability = txtLiability.Text
    Withheld = txtWithheld.Text
    
    Open App.Path & "\Store.txt" For Append As #1
    Write #1, N, AGIInput, Income, Liability, Withheld
    Close #1
    
End Sub
