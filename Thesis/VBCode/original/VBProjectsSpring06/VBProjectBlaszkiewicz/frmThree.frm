VERSION 5.00
Begin VB.Form frmThree 
   BackColor       =   &H80000003&
   Caption         =   "Buy a Car in 5 steps !"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtClient 
      Height          =   495
      Left            =   2400
      TabIndex        =   19
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdComparison 
      Caption         =   "Go to: Compare Cars"
      Height          =   1095
      Left            =   7320
      TabIndex        =   18
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   8640
      TabIndex        =   15
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back to Store"
      Height          =   1095
      Left            =   6120
      TabIndex        =   14
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   4560
      TabIndex        =   13
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "BUY"
      Height          =   615
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox txtCol 
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.PictureBox picBuy 
      Height          =   2775
      Left            =   5880
      ScaleHeight     =   2715
      ScaleWidth      =   3795
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for other models"
      Height          =   855
      Left            =   8400
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAvalon 
      Caption         =   "Avalon"
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdPrius 
      Caption         =   "Prius"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCamrySol 
      Caption         =   "Camry Solara"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCamry 
      Caption         =   "Camry"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdMatrix 
      Caption         =   "Matrix"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCorolla 
      Caption         =   "Corolla"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblClient 
      BackColor       =   &H8000000C&
      Caption         =   "3. ENTER YOUR NAME:"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblConfirm 
      BackColor       =   &H8000000C&
      Caption         =   "5. CONFIRM:"
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblBuy 
      BackColor       =   &H8000000C&
      Caption         =   "4. PURCHASE:"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblColor 
      BackColor       =   &H8000000C&
      Caption         =   "2. ENTER COLOR:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H8000000C&
      Caption         =   "1. SELECT A MODEL:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image imgCruiser 
      Height          =   9000
      Left            =   -240
      Picture         =   "frmThree.frx":0000
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmThree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Single
Dim Y As String
Dim Truck(1 To 10) As String
Dim Prices(1 To 10) As Single
Dim A, B As Integer
Dim total As Single


Private Sub cmdAvalon_Click()
    picBuy.Print "Model:", , "Base Price:"
    picBuy.Print Car(6), , FormatCurrency(Price(6))
    X = Price(6)
    
End Sub

Private Sub cmdBack_Click()
    frmThree.Visible = False
    frmOne.Visible = True
End Sub

Private Sub cmdBuy_Click()
    
   
    Dim Tax As Single
    
    Y = txtCol.Text
    picBuy.Print "------------------------------"
    picBuy.Print "Color:", , Y
    picBuy.Print "------------------------------"
    
    Tax = X * 0.07
    total = X + Tax
    
    picBuy.Print "Tax is:", , FormatCurrency(Tax)
    picBuy.Print "Total Price is:", , FormatCurrency(total)
End Sub

Private Sub cmdCamry_Click()
    picBuy.Print "Model:", , "Base Price:"
    picBuy.Print Car(3), , FormatCurrency(Price(3))
    X = Price(3)
    
End Sub

Private Sub cmdCamrySol_Click()
    picBuy.Print "Model:", , "Base Price:"
    picBuy.Print Car(4), , FormatCurrency(Price(4))
    X = Price(4)
    
End Sub

Private Sub cmdClear_Click()
    picBuy.Cls
    Y = False
End Sub

Private Sub cmdComparison_Click()
    frmThree.Visible = False
    frmTwo.Visible = True
End Sub

Private Sub cmdConfirm_Click()
    Dim C As String
    Dim pos As Integer
    Dim price2 As Single
    C = txtClient.Text
    If InStr(C, "a") <> 0 Then
        price2 = total * 0.9
        picBuy.Print "-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+"
        picBuy.Print "Congratulations "; C; "!!!"
        picBuy.Print "Your name contains letter 'a'"
        picBuy.Print "You are getting a 10% discount!!"
        picBuy.Print "The new price of your car is:", FormatCurrency(price2)
    Else
        MsgBox "Congratulations!! You are an owner of a new Toyota!!", , "Purchase Complete"
    End If
End Sub

Private Sub cmdCorolla_Click()
    picBuy.Print "Model:", , "Base Price:"
    picBuy.Print Car(1), , FormatCurrency(Price(1))
    X = Price(1)
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdMatrix_Click()
    picBuy.Print "Model:", , "Base Price:"
    picBuy.Print Car(2), , FormatCurrency(Price(2))
    X = Price(2)
    
End Sub

Private Sub cmdPrius_Click()
    picBuy.Print "Model:", , "Base Price:"
    picBuy.Print Car(5), , FormatCurrency(Price(5))
    X = Price(5)
    
End Sub

Private Sub cmdSearch_Click()
    
    Dim Found As Boolean
    Dim X As String
    Dim N As Integer
    
    X = InputBox("Enter type of car:", "Search for other models")
    
    Found = False
    Do While (Not Found) And (N < B)
        N = N + 1
        If Truck(N) = X Then
            Found = True
        End If
    Loop
    If (Found) Then
        picBuy.Print "Model:", , "Base Price:"
        picBuy.Print Truck(N), , FormatCurrency(Prices(N))
        picBuy.Print "***************************************"
        picBuy.Print "*"; "Sorry. This model is out of stock"; "*"
        picBuy.Print "***************************************"
    Else
         MsgBox "Model not found", , "Search Results"
    End If
    
End Sub

Private Sub Form_Load()
    Dim pos, Size As Integer
    pos = 0
    Open App.Path & "\CarData.txt" For Input As #2
    Do Until EOF(2)
        pos = pos + 1
        Input #2, Car(pos), Price(pos), MPG(pos)
    Loop
    Close #2
    Size = pos
    
    Open App.Path & "\CarData2.txt" For Input As #3
    Do Until EOF(3)
        A = A + 1
        Input #3, Truck(A), Prices(A)
    Loop
    Close #3
    B = A
End Sub




