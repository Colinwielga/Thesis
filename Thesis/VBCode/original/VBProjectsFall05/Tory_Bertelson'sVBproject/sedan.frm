VERSION 5.00
Begin VB.Form Sedan 
   BackColor       =   &H8000000D&
   Caption         =   "Sedan"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form5"
   ScaleHeight     =   8505
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find my sedan"
      Height          =   1575
      Left            =   2520
      TabIndex        =   4
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Click to go back to the main screen"
      Height          =   1695
      Left            =   1080
      TabIndex        =   3
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton cmdprice_Sedan 
      Caption         =   "Maximum Desired Price"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   3975
      Left            =   -240
      Picture         =   "sedan.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.PictureBox picresults_Sedan 
      Height          =   8535
      Left            =   4800
      ScaleHeight     =   8475
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "sedan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(FORM:CHOOSE YOUR CAR)
'TORY BERTELSON
'10-23-05
'THIS FORM ALLOWS THE USER TO NARROW DOWN HIS VEHICLE CHOICES DEPENDING ON PRICE

Option Explicit
Dim price As Single
Dim gas As Integer

Private Sub Cmdfind_Click()


Dim modelSedan(1 To 100) As String       'creates an array
Dim CTR As Integer
Dim priceSedan(1 To 100) As Single       'creates an array
Dim gasSedan(1 To 100) As Single         'creates an array
Dim J As Integer



    CTR = 0
    Open App.Path & "\Sedan.txt" For Input As #1       'opens the sports car file
        Do While Not EOF(1)
            CTR = CTR + 1                               'starts the counter
            Input #1, modelSedan(CTR), priceSedan(CTR), gasSedan(CTR)   'loads the three columns of data in to an array and names them
        Loop
        
            For J = 1 To CTR
                If price >= priceSedan(J) Then          'determines whether the price entered is equal to or greater than the prices in the file
                    picresults_Sedan.Print modelSedan(J), , FormatCurrency(priceSedan(J)), gasSedan(J)
                End If
    
            Next J
        
    Close #1
End Sub



Private Sub cmdmain_Click() 'allows the user to return to the main screen
Form1.Visible = True
sedan.Visible = False
Sport.Visible = False
SUV.Visible = False
Truck.Visible = False
End Sub

Private Sub cmdprice_Sedan_Click()
picresults_Sedan.Cls
picresults_Sedan.Print "Make and Model", "Price", "Reletive Gas Mileage"
picresults_Sedan.Print , , , ; "10 being Great and 1 being very bad"
picresults_Sedan.Print "____________________________________________________________________________"
price = InputBox("in thousands with no comma", "Maximum Desired Price to Spend")    'allows the user to input a maximum price
End Sub

