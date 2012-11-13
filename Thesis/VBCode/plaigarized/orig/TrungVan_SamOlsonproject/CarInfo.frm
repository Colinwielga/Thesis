VERSION 5.00
Begin VB.Form CarInfo 
   BackColor       =   &H8000000D&
   Caption         =   "Car Info"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000002&
      Caption         =   "All models and prices"
      Height          =   495
      Left            =   5880
      MaskColor       =   &H0000FF00&
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8040
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check out model && price"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.PictureBox picresult 
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   8955
      TabIndex        =   4
      Top             =   1800
      Width           =   9015
   End
   Begin VB.PictureBox Picresult2 
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "    Car price"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   $"CarInfo.frx":0000
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   8535
   End
   Begin VB.Label Label1 
      Caption         =   "This page is to help you have an idea of how your dream car will look like and how affordable it can be to you!"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "CarInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Dim model As String, found As Boolean, k As Integer, name As Integer
  
  Picresult2.Cls
  
  found = False
  model = InputBox("Enter the car model")
  For k = 1 To ctr
     If UCase(model) = carmodels(k) Then
        found = True
        picresult.Picture = LoadPicture(App.Path & "\images\" & picname(k))
        Picresult2.Print FormatCurrency(price(k))
     End If
  Next k
  
  If Not found Then
     MsgBox ("Please reenter the car model! It is written in the parentheses")
  End If
      
End Sub

Private Sub Command3_Click()
  Dim j As Integer
  
  picresult.Cls
  picresult.Picture = LoadPicture()
  
  picresult.Print Tab(5); "Car models"; Tab(20); "Car name"; Tab(50); "Prices"
  picresult.ForeColor = vbRed
  picresult.Print Tab(5); "***********************************************************************"
  For j = 1 To ctr
          
      picresult.Print Tab(5); carmodels(j); Tab(20); carname(j); Tab(50); FormatCurrency(price(j))
  Next j

End Sub

Private Sub Command2_Click()
   generalpage.Show
   CarInfo.Hide
End Sub
