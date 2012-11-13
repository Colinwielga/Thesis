VERSION 5.00
Begin VB.Form frmVenues 
   BackColor       =   &H8000000D&
   Caption         =   "Venues"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form2"
   ScaleHeight     =   7260
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   6135
      Left            =   4440
      ScaleHeight     =   6075
      ScaleWidth      =   5475
      TabIndex        =   5
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdStadium 
      Caption         =   "Stadium"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdCity 
      Caption         =   "City"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlaces 
      Caption         =   "List"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Main Menu"
      Height          =   735
      Left            =   6360
      TabIndex        =   1
      Top             =   6360
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   0
      Picture         =   "frmplaces.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmVenues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Location(1 To 100) As String
Dim Year(1 To 100) As Single
Dim Stadiumname(1 To 100) As String
Dim CTR As Single



'This funcion increments the CTR as well as the postion, and is looking through an array to find
'the text provided by the user in the input box
'used the boolean function of found=true or false
'it has 2 arrays
' it is the same formula used in frmGoal
Private Sub cmdCity_Click()
Dim Pos As Integer
    Dim found As Boolean
    Dim Cityinput As String
    Dim City(1 To 100) As String
    Dim Years(1 To 100) As Single
    
    Open App.Path & "\City.txt" For Input As #10
    picresults.Cls
    
    CTR = 0
    
    Do Until EOF(10)
        CTR = CTR + 1
        Input #10, City(CTR), Years(CTR)
    Loop
    Close #10
    
    Cityinput = InputBox("Name a City that you want to find", "Name")
    
    found = False
    Pos = 0
    
    Do While (Pos < CTR)
        Pos = Pos + 1
        If City(Pos) = Cityinput Then
        picresults.Print Cityinput; " Hosted the UEFA Champions League FInal in "; Years(Pos)
        found = True
        
        End If
    Loop
    
    If found = True Then
        
    Else
        MsgBox "Sorry this City has not hosted UEFA Champions League", , "Sorry"
    End If
    
    Close #10
End Sub

Private Sub cmdMenu_Click()
frmChampions.Show
frmVenues.Hide
End Sub
'increment CTR each time that it passes through the loop to move to the next postion in the array and then prints
Private Sub cmdPlaces_Click()
picresults.Cls

picresults.Print "Location", "Year"
    picresults.Print "*******************************************"
    
    Open App.Path & "\Ven.txt" For Input As #6
    
    Do While Not EOF(6)

        CTR = CTR + 1
        
        Input #6, Location(CTR), Year(CTR)
        picresults.Print Location(CTR); Tab(30); Year(CTR)
        
    Loop
    
  
    picresults.Print "*****************************************"
   
    Close #6
    
    End Sub
'This funcion increments the CTR as well as the postion, and is looking through an array to find
'the text provided by the user in the input box
'used the boolean function of found=true or false
'it has 2 arrays
' it is the same formula used in frmGoal
Private Sub cmdStadium_Click()
        Dim Pos As Integer
    Dim found As Boolean
    Dim Stadiumname As String
    Dim Stadium(1 To 100) As String
    Dim Years(1 To 100) As Single
    
    Open App.Path & "\ven.txt" For Input As #11
    
    picresults.Cls
    
    CTR = 0
    
    Do Until EOF(11)
        CTR = CTR + 1
        Input #11, Stadium(CTR), Years(CTR)
    Loop
    Close #11
    
    Stadiumname = InputBox("Name a Stadium that you want to find", "Name")
    
    found = False
    Pos = 0
    
    Do While (Pos < CTR)
        Pos = Pos + 1
        If Stadium(Pos) = Stadiumname Then
        picresults.Print Stadiumname; " Hosted the UEFA Champions League Final in "; Years(Pos)
        found = True
        
        End If
    Loop
    
    If found = True Then
        
    Else
        MsgBox "Sorry this Stadium has not hosted UEFA Champions League Final Game", , "Sorry"
    End If
    
    Close #11

End Sub


Private Sub cmdYear_Click()
Dim City(1 To 22) As String
    Dim Year(1 To 22) As Integer
    Dim Stadium As String
    Dim I As Integer
    Dim Pos As Integer
    Dim found As Boolean
    Dim size As Integer
    size = 22
    found = False
    Pos = 0
    'does the same as the City (Previous) funcion and it gives similar information as well.
        Do Until EOF(4)
            Pos = Pos + 1 'increments pos as it goes through the array
        Input #4, City(Pos), Year(Pos)
        Loop
    Close #3
    Pos = 0
    Stadium = InputBox("Name a Stadium", "Stadium")
        Do While found = False And Pos < size
            Pos = Pos + 1
            If Stadium = City(Pos) Then
                found = True
            End If
     Loop
        If found = True Then
            picresults.Print " The final of the UEFA Champions League was played in"; , City(Pos); "in the year"; Year(Pos)
            
            
        Else
            MsgBox "Sorry, There has been no Champions League Final in the Stadium you selected", , "Try Again!!"
        End If
End Sub
