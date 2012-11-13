VERSION 5.00
Begin VB.Form Truck 
   BackColor       =   &H8000000D&
   Caption         =   "Truck"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   ForeColor       =   &H80000016&
   LinkTopic       =   "Form4"
   ScaleHeight     =   8565
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find my truck"
      Height          =   1455
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Back to Main Screen"
      Height          =   1575
      Left            =   1440
      TabIndex        =   3
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrice_truck 
      BackColor       =   &H80000011&
      Caption         =   "maximum desired price"
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   960
      Picture         =   "Truck.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   3360
      Width           =   3615
      Begin VB.PictureBox Picture3 
         Height          =   3855
         Left            =   0
         Picture         =   "Truck.frx":735F6
         ScaleHeight     =   3795
         ScaleWidth      =   7635
         TabIndex        =   5
         Top             =   0
         Width           =   7695
      End
   End
   Begin VB.PictureBox picresults_Truck 
      Height          =   8295
      Left            =   5640
      ScaleHeight     =   8235
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Truck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(FORM:CHOOSE YOUR TRUCK)
'TORY BERTELSON
'10-23-05
'THIS FORM ALLOWS THE USER TO NARROW DOWN HIS VEHICLE CHOICES DEPENDING ON PRICE


Option Explicit
Dim price As Single
Dim gas As Integer

Private Sub Cmdfind_Click()


Dim modelTruck(1 To 100) As String       'creates an array
Dim CTR As Integer
Dim priceTruck(1 To 100) As Single       'creates an array
Dim gasTruck(1 To 100) As Single         'creates an array
Dim J As Integer



    CTR = 0
    Open App.Path & "\Truck.txt" For Input As #1       'opens the sports car file
        Do While Not EOF(1)
            CTR = CTR + 1                               'starts the counter
            Input #1, modelTruck(CTR), priceTruck(CTR), gasTruck(CTR)   'loads the three columns of data in to an array and names them
        Loop
        
            For J = 1 To CTR
                If price >= priceSport(J) Then          'determines whether the price entered is equal to or greater than the prices in the file
                    picresults_Truck.Print modelTruck(J), , FormatCurrency(priceTruck(J)), gasTruck(J)
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

Private Sub cmdPrice_truck_Click()

picresults_Truck.Cls

picresults_Truck.Print "Make and Model", "Price", "Reletive Gas Mileage"
picresults_Truck.Print , , , ; "10 being Great and 1 being very bad"
picresults_Truck.Print "____________________________________________________________________________"
price = InputBox("in thousands with no comma", "Maximum Desired Price to Spend")    'allows the user to imput a maximum desired price
End Sub

