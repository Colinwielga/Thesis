VERSION 5.00
Begin VB.Form frmSantorini 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   Picture         =   "frmSantorini.frx":0000
   ScaleHeight     =   8205
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Height          =   375
      Left            =   5880
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtNumDays 
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtHotelName 
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   6000
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
      Left            =   3360
      TabIndex        =   9
      Top             =   5040
      Width           =   2175
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
      Height          =   615
      Left            =   5880
      TabIndex        =   8
      Top             =   6480
      Width           =   2175
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
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   6480
      Width           =   2415
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
      Left            =   4800
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
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
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   1935
      Left            =   1080
      ScaleHeight     =   1875
      ScaleWidth      =   6315
      TabIndex        =   3
      Top             =   3000
      Width           =   6375
   End
   Begin VB.CommandButton cmdBacktoHome 
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
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdSantoriniInfo 
      Caption         =   "Info on Santorini"
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
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
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
      Left            =   5880
      TabIndex        =   15
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lblNumDays 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
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
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lblHotelName 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
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
      Left            =   840
      TabIndex        =   11
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label lblSort 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
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
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblSantorini 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Santorini"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmSantorini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HotelName(1 To 100) As String, HotelPrice(1 To 100) As Single, City(1 To 100) As String, CTR As Integer, Pass As Integer, Pos As Integer, Temp As String, Temp2 As Single, I As Integer
'Project Name: Ideal Greek Island
'Form: frmSantorini
'Author: Alie Chandler
'Date Writen: started 3/16 finished 3/23
'Form Objective: This form will provide information about the island for the user, give hotel information (alphabetically and by price), allow the user to find the total hotel cost of their trip, and give them the option to book their hotel or search other islands.
Private Sub cmdAlpha_Click()
    CTR = 0
    Open App.Path & "\SantoriniHotels.txt" For Input As #7
    
    Do Until EOF(7)
        CTR = CTR + 1
        Input #7, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #7
    
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

Private Sub cmdBacktoHome_Click()
    frmSantorini.Hide
    frmHome1.Show
End Sub

Private Sub cmdBook_Click()
    frmSantorini.Hide
    frmReserveHotel.Show
End Sub

Private Sub cmdCheckCost_Click()
    Dim Found As Boolean, I As Integer, Days As Integer, HName As String, Total As Single, Tax As Single, FinalTotal As Single
    CTR = 0
    Open App.Path & "\SantoriniHotels.txt" For Input As #7
    
    Do Until EOF(7)
        CTR = CTR + 1
        Input #7, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #7
    
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
    Open App.Path & "\SantoriniHotels.txt" For Input As #7
    
    Do Until EOF(7)
        CTR = CTR + 1
        Input #7, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #7
    
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



Private Sub cmdSantoriniInfo_Click()
    MsgBox "Santorini is one of the most magical islands of Greece. It is a barren, rocky island just opposite a volcano, with black and red beaches and towns situated on high cliffs offering breathtaking views and fantastic sunsets. Santorini has a dramatic beauty as opposed to lush and green islands. If you can, you should try to stay in Fira, Imerovigli or Oia, the towns on the cliffs, which are very beautiful and full of little cafes, shops and places of interest. There is a bus that goes to the beaches every day, and it is much better to be in the towns in the evening and on the beaches during the day. If you stay in Monolithos you will have more peace and quiet.(Taken from:http://www.in2greece.com/)", , "Santorini Information"
End Sub

