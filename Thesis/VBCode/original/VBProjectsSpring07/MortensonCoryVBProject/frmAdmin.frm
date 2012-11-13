VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00000000&
   Caption         =   "Administration Page"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form2"
   ScaleHeight     =   7305
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H0000FFFF&
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H0000FFFF&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtPerson 
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   5655
      Left            =   4800
      ScaleHeight     =   5595
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   "Client's Name:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR2 As Integer




Private Sub cmdDisplay_Click()
    Dim Found As Boolean, Pos As Integer
    Dim Person As String
    
    'This searches the Client Array performing a match and stop search.
    picResults.Cls
    
    Pos = 0
    Found = False
    Person = txtPerson.Text
    
    Do Until Found = True Or Pos > CTR2
        Pos = Pos + 1
        If Person = ClientArray(Pos) Then
            Found = True
        End If
    Loop
    
    If Found = True Then
        picResults.Print ClientArray(Pos)
        picResults.Print "Adjusted Gross Income ", FormatCurrency(AGIArray(Pos))
        picResults.Print "Taxable Income ", FormatCurrency(IncomeArray(Pos))
        picResults.Print "Tax Liability", Tab(20), FormatCurrency(LiabilityArray(Pos))
        picResults.Print "Taxes Withheld", FormatCurrency(WithheldArray(Pos))
    Else
        MsgBox "No client exists with the name " & Person, , "Error"
    End If
    
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdLoad_Click()
    'This loads the previously stored infomration into arrays.
    
    Open App.Path & "\Store.txt" For Input As #1
    CTR2 = 0
    
    Do Until EOF(1)
        CTR2 = CTR2 + 1
        Input #1, ClientArray(CTR2), AGIArray(CTR2), IncomeArray(CTR2), LiabilityArray(CTR2), WithheldArray(CTR2)
    Loop
    Close #1
    
    'Open App.Path & "\Store.txt" For Append As #1
    'Write #1,
    'Close #1
End Sub

Private Sub cmdSort_Click()
    Dim Pass As Integer, Pos As Integer, Temp As String, A As Integer
    Dim Temp1 As Double, Temp2 As Double, Temp3 As Double, Temp4 As Double
    picResults.Cls
    
    'Uses the bubble sort to arrange the clients in descending order according to AGI
    For Pass = 1 To (CTR2 - 1)
        For Pos = 1 To (CTR2 - Pass)
            If AGIArray(Pos) < AGIArray(Pos + 1) Then
                Temp = ClientArray(Pos)
                ClientArray(Pos) = ClientArray(Pos + 1)
                ClientArray(Pos + 1) = Temp
            
                Temp1 = AGIArray(Pos)
                AGIArray(Pos) = AGIArray(Pos + 1)
                AGIArray(Pos + 1) = Temp1
                
                Temp2 = IncomeArray(Pos)
                IncomeArray(Pos) = IncomeArray(Pos + 1)
                IncomeArray(Pos + 1) = Temp2
                
                Temp3 = LiabilityArray(Pos)
                LiabilityArray(Pos) = LiabilityArray(Pos + 1)
                LiabilityArray(Pos + 1) = Temp3
                
                Temp4 = WithheldArray(Pos)
                WithheldArray(Pos) = WithheldArray(Pos + 1)
                WithheldArray(Pos + 1) = Temp4
                
            End If
        Next Pos
    Next Pass
    
    'Print the list that was just created
    For A = 1 To CTR2
        picResults.Print ClientArray(A), Tab(2); FormatCurrency(AGIArray(A))
    Next A
    
End Sub
