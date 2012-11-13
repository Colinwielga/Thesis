VERSION 5.00
Begin VB.Form frmMykonos 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   Picture         =   "frmMykonos.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Height          =   375
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   14
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox txtNumDays 
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtHotelName 
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   5640
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
      Left            =   3120
      TabIndex        =   9
      Top             =   4680
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
      Left            =   4920
      TabIndex        =   8
      Top             =   6120
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
      Left            =   1080
      TabIndex        =   6
      Top             =   6120
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
      Left            =   4560
      TabIndex        =   5
      Top             =   2160
      Width           =   2295
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   960
      ScaleHeight     =   1755
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   2760
      Width           =   6255
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
      Height          =   735
      Left            =   5400
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdMykonosInfo 
      Caption         =   "Info on Mykonos"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label lblNumDays 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      Left            =   3480
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblHotelName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblSort 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      Left            =   3000
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblMykonos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mykonos"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMykonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HotelName(1 To 100) As String, HotelPrice(1 To 100) As Single, City(1 To 100) As String, CTR As Integer, Pass As Integer, Pos As Integer, Temp As String, Temp2 As Single, I As Integer
'Project Name: Ideal Greek Island
'Form: frmMykonos
'Author: Alie Chandler
'Date Writen: started 3/16 finished 3/23
'Form Objective: This form will provide information about the island for the user, give hotel information (alphabetically and by price), allow the user to find the total hotel cost of their trip, and give them the option to book their hotel or search other islands.
Private Sub cmdAlpha_Click()
    CTR = 0
    Open App.Path & "\MykonosHotels.txt" For Input As #4
    
    Do Until EOF(4)
        CTR = CTR + 1
        Input #4, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #4
    
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
    frmMykonos.Hide
    frmHome1.Show
End Sub

Private Sub cmdBook_Click()
    frmMykonos.Hide
    frmReserveHotel.Show
End Sub

Private Sub cmdCheckCost_Click()
    Dim Found As Boolean, I As Integer, Days As Integer, HName As String, Total As Single, Tax As Single, FinalTotal As Single
    CTR = 0
    Open App.Path & "\MykonosHotels.txt" For Input As #4
    
    Do Until EOF(4)
        CTR = CTR + 1
        Input #4, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #4
    
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

Private Sub cmdMykonosInfo_Click()
    MsgBox "Mykonos is a grand example of unique Cycladic architecture set around a picturesque fishing-village bay. Totally whitewashed organic cube-like buildings fit closely together to form a kind of haphazard maze of narrow alley ways and streets. The earthen colors of the bare hills which surround the town's gleaming whiteness is set between the aura of an incredibly blue sky and even deeper blue sparkling sea. Together with being friendly and open people, the locals have a healthy understanding of what it means to have a good time. Put this together with all the island's other qualities and it is no wonder Mykonos has been often named the jewel of the Aegean Sea.  In relation to the rest of Greece, Mykonos can be one of the more expensive places to visit.(Taken from:http://www.mykonos-web.com/)", , "Mykonos Information"
End Sub

Private Sub cmdPrice_Click()
    CTR = 0
    Open App.Path & "\MykonosHotels.txt" For Input As #4
    
    Do Until EOF(4)
        CTR = CTR + 1
        Input #4, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #4
    
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


