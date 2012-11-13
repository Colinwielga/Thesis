VERSION 5.00
Begin VB.Form SUV 
   BackColor       =   &H8000000D&
   Caption         =   "Sport Utility Vehicle"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form3"
   ScaleHeight     =   8550
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find the car for you"
      Height          =   1335
      Left            =   2880
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "click to go back to the main screen"
      Height          =   1575
      Left            =   1080
      TabIndex        =   3
      Top             =   6720
      Width           =   2895
   End
   Begin VB.CommandButton cmdprice_SUV 
      Caption         =   "maximum desired price"
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   3975
      Left            =   0
      Picture         =   "SUV.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.PictureBox picresults_SUV 
      Height          =   8775
      Left            =   5040
      ScaleHeight     =   8715
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "SUV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(FORM:CHOOSE YOUR SUV)
'TORY BERTELSON
'10-23-05
'THIS FORM ALLOWS THE USER TO NARROW DOWN HIS VEHICLE CHOICES DEPENDING ON PRICE


Option Explicit
Dim price As Single
Dim gas As Integer

Private Sub Cmdfind_Click()


Dim modelSUV(1 To 100) As String       'creates an array
Dim CTR As Integer
Dim priceSUV(1 To 100) As Single       'creates an array
Dim gasSUV(1 To 100) As Single         'creates an array
Dim J As Integer



    CTR = 0
    Open App.Path & "\SUV.txt" For Input As #1       'opens the sports car file
        Do While Not EOF(1)
            CTR = CTR + 1                               'starts the counter
            Input #1, modelSUV(CTR), priceSUV(CTR), gasSUV(CTR)   'loads the three columns of data in to an array and names them
        Loop
        
            For J = 1 To CTR
                If price >= priceSUV(J) Then          'determines whether the price entered is equal to or greater than the prices in the file
                    picresults_SUV.Print modelSUV(J), , FormatCurrency(priceSUV(J)), gasSUV(J)
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

Private Sub cmdprice_SUV_Click()
picresults_SUV.Cls

picresults_SUV.Print "Make and Model", "Price", "Reletive Gas Mileage"
picresults_SUV.Print , , , ; "10 being Great and 1 being very bad"
picresults_SUV.Print "____________________________________________________________________________"
price = InputBox("in thousands with no comma", "Maximum Desired Price to Spend")    'allows the user to input a maximum desired price


End Sub
