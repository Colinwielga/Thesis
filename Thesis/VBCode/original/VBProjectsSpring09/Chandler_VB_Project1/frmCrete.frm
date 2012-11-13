VERSION 5.00
Begin VB.Form frmCrete 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmCrete.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCostCheck 
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
      Left            =   3840
      TabIndex        =   15
      Top             =   5040
      Width           =   1935
   End
   Begin VB.PictureBox picResults2 
      Height          =   495
      Left            =   6480
      ScaleHeight     =   435
      ScaleWidth      =   2115
      TabIndex        =   13
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox txtNumDays 
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox txtHotelName 
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   6000
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
      Left            =   5280
      TabIndex        =   8
      Top             =   6720
      Width           =   1575
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
      Left            =   2760
      TabIndex        =   7
      Top             =   6720
      Width           =   1575
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
      Height          =   615
      Left            =   5160
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
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
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   1935
      Left            =   1440
      ScaleHeight     =   1875
      ScaleWidth      =   6915
      TabIndex        =   3
      Top             =   3000
      Width           =   6975
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
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreteInfo 
      Caption         =   "Info About Crete"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   6600
      TabIndex        =   14
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblNumDays 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   3960
      TabIndex        =   12
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblHotelName 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Left            =   1440
      TabIndex        =   10
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblSort 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Find Hotels:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblCrete 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Crete"
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
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmCrete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HotelName(1 To 100) As String, HotelPrice(1 To 100) As Single, City(1 To 100) As String, CTR As Integer, Pass As Integer, Pos As Integer, Temp As String, Temp2 As Single, I As Integer
'Project Name: Ideal Greek Island
'Form: frmCrete
'Author: Alie Chandler
'Date Writen: started 3/16 finished 3/23
'Form Objective: This form will provide information about the island for the user, give hotel information (alphabetically and by price), allow the user to find the total hotel cost of their trip, and give them the option to book their hotel or search other islands.
Private Sub cmdAlpha_Click()
    CTR = 0
    Open App.Path & "\CreteHotels.txt" For Input As #2
    
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #2
    
    picResults.Cls
    'sort hotel array by name(same for all island forms)
    picResults.Print "Hotel Name"; Tab(20); "Price (Double Room/night)"; Tab(50); "City"
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
        picResults.Print HotelName(I); Tab(20); FormatCurrency(HotelPrice(I)); Tab(50); City(I)
    Next I
End Sub

Private Sub cmdBacktoHome_Click()
    frmCrete.Hide
    frmHome1.Show
End Sub

Private Sub cmdBook_Click()
    frmCrete.Hide
    frmReserveHotel.Show
End Sub

Private Sub cmdCostCheck_Click()
    Dim Found As Boolean, I As Integer, Days As Integer, HName As String, Total As Single, Tax As Single, FinalTotal As Single
    
    CTR = 0
    Open App.Path & "\CreteHotels.txt" For Input As #2
    
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #2
    'calculates base hotel price * days stay *7% tax (same for all island forms)
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

Private Sub cmdCreteInfo_Click()
    'give info about island (same for all island forms)
    MsgBox "From fertile coastal plains to rugged barren mountains, from mellow stone houses to stark concrete modernity, from bustling capital to sleepy hill villages, Crete, the largest of the Greek Islands, is an island of contrasts. Home to around 650,000 people and several million olive trees, the island remains ever popular with visitors from northern Europe, other parts of Greece, and indeed, visitors from all over the World. It consists of four prefectures: Chania, Rethymnon, Heraklion and Lasithi. (Taken from: http://www.ellada.net/crete-info/)", , "Crete Information"
End Sub

Private Sub cmdPrice_Click()
    
    CTR = 0
    Open App.Path & "\CreteHotels.txt" For Input As #2
    
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #2
    
    picResults.Cls
    'sort hotel array by price (same for all island forms)
    picResults.Print "Hotel Name"; Tab(20); "Price (Double Room/night)"; Tab(50); "City"
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
        picResults.Print HotelName(I); Tab(20); FormatCurrency(HotelPrice(I)); Tab(50); City(I)
    Next I
End Sub

Private Sub cmdQuit_Click()
    End
End Sub


