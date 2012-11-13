VERSION 5.00
Begin VB.Form frmHydra 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   Picture         =   "frmHydra.frx":0000
   ScaleHeight     =   8640
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   14
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox txtNumDays 
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox txtHotelName 
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheckCost 
      BackColor       =   &H00FFFFFF&
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
      Height          =   615
      Left            =   3480
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
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   6960
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
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   6960
      Width           =   2295
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
      Width           =   1695
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
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   1560
      ScaleHeight     =   1755
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   3120
      Width           =   5895
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
      Left            =   6240
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdHydraInfo 
      Caption         =   "Info on Hydra"
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
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblNumDays 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Left            =   3600
      TabIndex        =   13
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label lblHotelName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Left            =   1200
      TabIndex        =   11
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label lblSort 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Left            =   3600
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblHydra 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Hydra"
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
      Width           =   2415
   End
End
Attribute VB_Name = "frmHydra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HotelName(1 To 100) As String, HotelPrice(1 To 100) As Single, City(1 To 100) As String, CTR As Integer, Pass As Integer, Pos As Integer, Temp As String, Temp2 As Single, I As Integer
'Project Name: Ideal Greek Island
'Form: frmHydra
'Author: Alie Chandler
'Date Writen: started 3/16 finished 3/23
'Form Objective: This form will provide information about the island for the user, give hotel information (alphabetically and by price), allow the user to find the total hotel cost of their trip, and give them the option to book their hotel or search other islands.
Private Sub cmdAlpha_Click()
    CTR = 0
    Open App.Path & "\HydraHotels.txt" For Input As #3
    
    Do Until EOF(3)
        CTR = CTR + 1
        Input #3, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #3
    
    picResults.Cls
    
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
    frmHydra.Hide
    frmHome1.Show
End Sub

Private Sub cmdBook_Click()
    frmHydra.Hide
    frmReserveHotel.Show
End Sub

Private Sub cmdCheckCost_Click()
    Dim Found As Boolean, I As Integer, Days As Integer, HName As String, Total As Single, Tax As Single, FinalTotal As Single
    CTR = 0
    Open App.Path & "\HydraHotels.txt" For Input As #3
    
    Do Until EOF(3)
        CTR = CTR + 1
        Input #3, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #3
    
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

Private Sub cmdHydraInfo_Click()
    MsgBox "Hydra Island Greece, one of the most un-spoiled and interesting of the Greek islands, is a small rocky island in the Argo Saronic Gulf, south east out of the Athens port of Piraeus and within sight of the southern Peloponnese mainland. It's very cosmopolitan, safe and one of the easiest of the Greek islands to get to. Best of all, the entire island is a preserved national monument and has retained all its 17th & 18th century charm and quaintness. (Taken from:http://www.hydradirect.com/about_hydra)", , "Hydra Information"
End Sub

Private Sub cmdPrice_Click()
    CTR = 0
    Open App.Path & "\HydraHotels.txt" For Input As #3
    
    Do Until EOF(3)
        CTR = CTR + 1
        Input #3, HotelName(CTR), HotelPrice(CTR), City(CTR)
    Loop
    Close #3
    
    picResults.Cls
    
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

