VERSION 5.00
Begin VB.Form Sport 
   BackColor       =   &H8000000D&
   Caption         =   "Sports Cars"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8730
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdfind 
      Caption         =   "find sports car"
      Height          =   1455
      Left            =   2520
      TabIndex        =   4
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Click to go back to the main screen"
      Height          =   1695
      Left            =   960
      TabIndex        =   3
      Top             =   6360
      Width           =   2895
   End
   Begin VB.PictureBox picresults_Sport 
      Height          =   8655
      Left            =   5040
      ScaleHeight     =   8595
      ScaleWidth      =   5715
      TabIndex        =   2
      Top             =   0
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   -120
      Picture         =   "Sport.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   -480
      Width           =   6015
   End
   Begin VB.CommandButton Cmdprice_sport 
      Caption         =   "Maximum Desired Price"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   2295
   End
End
Attribute VB_Name = "Sport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(FORM:CHOOSE YOUR SPORTS CAR)
'TORY BERTELSON
'10-23-05
'THIS FORM ALLOWS THE USER TO NARROW DOWN HIS VEHICLE CHOICES DEPENDING ON PRICE

Option Explicit
Dim price As Single
Dim gas As Integer

Private Sub Cmdfind_Click()


Dim modelSport(1 To 100) As String       'creates an array
Dim CTR As Integer
Dim priceSport(1 To 100) As Single       'creates an array
Dim gasSport(1 To 100) As Single         'creates an array
Dim J As Integer



    CTR = 0
    Open App.Path & "\sports.txt" For Input As #1       'opens the sports car file
        Do While Not EOF(1)
            CTR = CTR + 1                               'starts the counter
            Input #1, modelSport(CTR), priceSport(CTR), gasSport(CTR)   'loads the three columns of data in to an array and names them
        Loop
        
            For J = 1 To CTR
                If price >= priceSport(J) Then          'determines whether the price entered is equal to or greater than the prices in the file
                    picresults_Sport.Print modelSport(J), , FormatCurrency(priceSport(J)), gasSport(J)
                End If
    
            Next J
        
    Close #1
End Sub

Private Sub cmdmain_Click() 'allows the user to go back to the main screen
Form1.Visible = True
sedan.Visible = False
Sport.Visible = False
SUV.Visible = False
Truck.Visible = False
End Sub




Private Sub Cmdprice_sport_Click()

picresults_Sport.Cls        'clears any data in the picture box

picresults_Sport.Print "Make and Model", "Price", "Reletive Gas Mileage"
picresults_Sport.Print , , , ; "10 being Great and 1 being very bad"
picresults_Sport.Print "____________________________________________________________________________"
price = InputBox("in thousands with no comma", "Maximum Desired Price to Spend")    'allows the user to input a desired maximum price


End Sub

Private Sub Command1_Click()

End Sub


