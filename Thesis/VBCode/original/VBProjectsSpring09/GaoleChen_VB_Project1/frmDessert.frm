VERSION 5.00
Begin VB.Form frmDessert 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   Picture         =   "frmDessert.frx":0000
   ScaleHeight     =   8535
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order"
      Height          =   375
      Left            =   11520
      TabIndex        =   10
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Height          =   375
      Left            =   11520
      TabIndex        =   8
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdNiangao 
      Caption         =   "Niangao"
      Height          =   855
      Left            =   3600
      TabIndex        =   7
      Top             =   7320
      Width           =   1455
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00C0FFFF&
      Height          =   2895
      Left            =   10680
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuti 
      Caption         =   "Quit"
      Height          =   855
      Left            =   9360
      TabIndex        =   4
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   855
      Left            =   7440
      TabIndex        =   3
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Get the Dessert Menu"
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   7320
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0FFFF&
      Height          =   3615
      Left            =   9000
      ScaleHeight     =   3555
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter the number to see the picture"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      TabIndex        =   9
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "==>Just Niangao"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
End
Attribute VB_Name = "frmDessert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Digital Menu
'Form Name: frmDessert
'Authors: Gaole Chen
'Date Written: 3/16/09
'Objective: Here the user can order desserts. The picturebox will also give a picture of each one.

Option Explicit
Dim Dessert(1 To 15) As String, Price(1 To 15) As Integer, CTR As Integer, runningTotal As Integer

Private Sub cmdBack_Click()
frmDessert.Hide
frmMain.Show
End Sub

Private Sub cmdEnter_Click()
'here the user can check out the pictures of the dessert
'declare the variables
Dim Number As Integer
Number = InputBox("Please enter a number from the menu.")
picResult.Picture = LoadPicture(App.Path & "\" & Number & ".jpg")
End Sub

Private Sub cmdNext_Click()
frmBeverage.Show
frmDessert.Hide
End Sub

Private Sub cmdNiangao_Click()
'declare the variables
Dim J As Integer, space(1 To 15) As Integer, Total As Integer
Total = 0
'clear the screen first
picResults.Cls
picResults.Print "All the Niangaos are here!"
For J = 1 To CTR
    If InStr(Dessert(J), "niangao") Then
        space(J) = InStr(Dessert(J), " ")
        picResults.Print Left(Dessert(J), space(J)); Tab(25); FormatCurrency(Price(J))
        Total = Total + Price(J)
    End If
Next J
picResults.Print
picResults.Print "They are all delicious!"
picResults.Print "And they only cost "; FormatCurrency(Total); " even if you order all of them!"
End Sub

Private Sub cmdOrder_Click()
Dim I As Integer
runningTotal = 0
picResults.Print "You want to order:"
I = InputBox("Please enter the number you wish to order.(Input 0 to indicate the end of order)")
Do While I <> 0
    picResults.Print Dessert(I); Tab(25); FormatCurrency(Price(I))
    runningTotal = runningTotal + Price(I)
    I = InputBox("Please enter the number you wish to order.(Input 0 to indicate the end of order)")
Loop
Totaldessertcost = runningTotal * 1.08
End Sub

Private Sub cmdQuti_Click()
End
End Sub

Private Sub cmdRead_Click()
'get the file
Open App.Path & "\Dessert.txt" For Input As #1
CTR = 0
Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Dessert(CTR), Price(CTR)
        
        'list all of the desserts and their price
        picResults.Print Dessert(CTR); Tab(25); FormatCurrency(Price(CTR)), CTR
              
    Loop
Close #1


End Sub

