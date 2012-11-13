VERSION 5.00
Begin VB.Form frmRhodes 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   Picture         =   "frmRhodes.frx":0000
   ScaleHeight     =   9270
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   14
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox txtNumDays 
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox txtHotelName 
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheckCost 
      Caption         =   "Check the Cost"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdBook 
      Caption         =   "Reserve Hotel Now!"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrice 
      Caption         =   "Sort By Price"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Sort Alphabetically"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   1200
      ScaleHeight     =   1755
      ScaleWidth      =   6075
      TabIndex        =   3
      Top             =   3000
      Width           =   6135
   End
   Begin VB.CommandButton cmdBackHome 
      Caption         =   "Search Other Islands"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdRhodesInfo 
      Caption         =   "Info on Rhodes"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Total Cost (7% tax)"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblNumDays 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Number of Days"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblHotelName 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter Hotel Name"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lblSort 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Find Hotels:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblRhodes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Rhodes"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmRhodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HotelName(1 To 100) As String, HotelPrice(1 To 100) As Single, City(1 To 100) As String, CTR As Integer, Pass As Integer, Pos As Integer, Temp As String, Temp2 As Single, I As Integer
'Project Name: Ideal Greek Island
'Form: frmRhodes
'Author: Alie Chandler
'Date Writen: started 3/16 finished 3/23
'Form Objective: This form will provide information about the island for the user, give hotel information (alphabetically and by price), allow the user to find the total hotel cost of their trip, and give them the option to book their hotel or search other islands.
Private Sub cmdAlpha_Click()
    CTR = 0
    Open App.Path & "\RhodesHotels.txt" For Input As #5
    
    Do Until EOF(5)
        CTR = CTR + 1
        Input #5, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #5
    
    picResults.Cls
    
    picResults.Print "Hotel Name"; Tab(30); "Price (Double Room/night)"; Tab(60); "City"
    picResults.Print "_________________________________________________________"
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If HotelName(Pos) > HotelName(Pos + 1) Then
                Temp = HotelName(Pos)
                HotelName(Pos) = HotelName(Pos + 1)
                HotelName(Pos + 1) = Temp
                Temp2 = HotelPrice(Pos)
                HotelPrice(Pos) = HotelPrice(Pos + 1)
                HotelPrice(Pos + 1) = Temp2
                Temp = City(Pos)
                City(Pos) = City(Pos + 1)
                City(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
                
    For I = 1 To CTR
        picResults.Print HotelName(I); Tab(30); FormatCurrency(HotelPrice(I)); Tab(60); City(I)
    Next I
End Sub

Private Sub cmdBackHome_Click()
    frmRhodes.Hide
    frmHome1.Show
End Sub

Private Sub cmdBook_Click()
    frmRhodes.Hide
    frmReserveHotel.Show
End Sub

Private Sub cmdCheckCost_Click()
    Dim Found As Boolean, I As Integer, Days As Integer, HName As String, Total As Single, Tax As Single, FinalTotal As Single
    CTR = 0
    Open App.Path & "\RhodesHotels.txt" For Input As #5
    
    Do Until EOF(5)
        CTR = CTR + 1
        Input #5, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #5
    
    HName = txtHotelName.Text
    Days = txtNumDays.Text
    I = 0
    Found = False
     
    Do While ((Not Found) And (I < CTR))
        I = I + 1
        If HName = HotelName(I) Then
            Found = True
        End If
    Loop
           
    If (Not Found) Then
        MsgBox "Hotel does not exist. Please read list above to find an available hotel.", , "Error"
        Else
        Total = Days * HotelPrice(I)
        Tax = Total * 0.07
        FinalTotal = Total + Tax
        picResults2.Cls
        picResults2.Print FormatCurrency(FinalTotal)
    End If
End Sub

Private Sub cmdPrice_Click()
    CTR = 0
    Open App.Path & "\RhodesHotels.txt" For Input As #5
    
    Do Until EOF(5)
        CTR = CTR + 1
        Input #5, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #5
    
    picResults.Cls
    
    picResults.Print "Hotel Name"; Tab(30); "Price (Double Room/night)"; Tab(60); "City"
    picResults.Print "_________________________________________________________"
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If HotelPrice(Pos) > HotelPrice(Pos + 1) Then
                Temp = HotelName(Pos)
                HotelName(Pos) = HotelName(Pos + 1)
                HotelName(Pos + 1) = Temp
                Temp2 = HotelPrice(Pos)
                HotelPrice(Pos) = HotelPrice(Pos + 1)
                HotelPrice(Pos + 1) = Temp2
                Temp = City(Pos)
                City(Pos) = City(Pos + 1)
                City(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
                
    For I = 1 To CTR
        picResults.Print HotelName(I); Tab(30); FormatCurrency(HotelPrice(I)); Tab(60); City(I)
    Next I
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRhodesInfo_Click()
    MsgBox "Rhodes is the largest island in the Dodecanese. Its capital city, located at its northern tip, is the capital of the Prefecture with the Medieval Town in its centre. In 1988 the Medieval Town was designated as a World Heritage City. The Medieval Town of Rhodes is the result of different architectures belonging to various historic eras, predominantly those of the Knights of St. John. The island of Rhodes is located at the crossroads of two major sea routes of the Mediterranean between the Aegean Sea and the coast of the Middle East, as well as Cyprus and Egypt. The meeting point of three continents, it has known many civilizations. Throughout its long history the different people who settled on Rhodes left their mark in all aspects of the island's culture: art, language, architecture. Its strategic position brought to the island great wealth and made the city of Rhodes one of the leading cities of the ancient Greek world. (Taken from:http://www.colossus.gr/information.htm)", , "Rhodes Information"
End Sub
