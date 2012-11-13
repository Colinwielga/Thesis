VERSION 5.00
Begin VB.Form frmSamos 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   Picture         =   "frmSamos.frx":0000
   ScaleHeight     =   7410
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtNumDays 
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtHotelName 
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   6120
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
      Left            =   3840
      TabIndex        =   9
      Top             =   5160
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
      Width           =   2535
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
      Left            =   1200
      TabIndex        =   6
      Top             =   6720
      Width           =   2775
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
      Left            =   5400
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
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   1935
      Left            =   1800
      ScaleHeight     =   1875
      ScaleWidth      =   6075
      TabIndex        =   3
      Top             =   3120
      Width           =   6135
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
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdSamosInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Info on Samos"
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
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      Left            =   6120
      TabIndex        =   15
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lblNumDays 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      Left            =   3720
      TabIndex        =   13
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lblHotelName 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      Left            =   1320
      TabIndex        =   11
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label lblSort 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      Left            =   3480
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblSamos 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Samos"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmSamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HotelName(1 To 100) As String, HotelPrice(1 To 100) As Single, City(1 To 100) As String, CTR As Integer, Pass As Integer, Pos As Integer, Temp As String, Temp2 As Single, I As Integer
'Project Name: Ideal Greek Island
'Form: frmSamos
'Author: Alie Chandler
'Date Writen: started 3/16 finished 3/23
'Form Objective: This form will provide information about the island for the user, give hotel information (alphabetically and by price), allow the user to find the total hotel cost of their trip, and give them the option to book their hotel or search other islands.
Private Sub cmdAlpha_Click()
    CTR = 0
    Open App.Path & "\SamosHotels.txt" For Input As #6
    
    Do Until EOF(6)
        CTR = CTR + 1
        Input #6, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #6
    
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
    frmSamos.Hide
    frmHome1.Show
End Sub

Private Sub cmdBook_Click()
    frmSamos.Hide
    frmReserveHotel.Show
End Sub

Private Sub cmdCheckCost_Click()
    Dim Found As Boolean, I As Integer, Days As Integer, HName As String, Total As Single, Tax As Single, FinalTotal As Single
    CTR = 0
    Open App.Path & "\SamosHotels.txt" For Input As #6
    
    Do Until EOF(6)
        CTR = CTR + 1
        Input #6, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #6
    
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
    Open App.Path & "\SamosHotels.txt" For Input As #6
    
    Do Until EOF(6)
        CTR = CTR + 1
        Input #6, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #6
    
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

Private Sub cmdSamosInfo_Click()
    MsgBox "Samos is one of the biggest islands of Greece and it's the favorite touristic destination in the north east aegean. It features a rich history as well as beautiful nature, an intact infrastructure and friendly people who will make your stay a memorable one, be it for one week or many years.  Samos lies on the eastern border of Greece and is less than 2 kilometers away from the turkish coastline, making it ideal for a trip abroad as well. It has around 34.000 inhabitants living in the four municipalities of Vathi (Samos), Karlovassi, Marathokambos and Pythagorio. (Taken from:http://www.samosinfo.com/)", , "Samos Information"
End Sub
